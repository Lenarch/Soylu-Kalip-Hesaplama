"""
Kapasite uygulamasının SQLite veritabanı.

Tablolar
--------
aktarimlar   : İşlem kaydı. Her Excel yüklemesi (dosya + ay) bir satır. Geri alınabilir.
siparisler   : Ürün + firma bazında haftalık sipariş adetleri (yıl + hafta + ay ile)
backlog      : Ürün + firma bazında, AY AY backlog ve stok adetleri
operasyonlar : Ürünün operasyonları, presi, alternatifleri, parça başı saniye
urunler      : (İsteğe bağlı) ürün açıklaması

Kural (her yıl + ay için):
  - 'yerine' yükleme: o ayın önceki yüklemelerinin yerine geçer (aynı dosyanın güncel hali).
  - 'ekle' yükleme  : o ayın üstüne eklenir (ör. başka müşterinin dosyası).
  Geçerli = en son aktif 'yerine' yükleme + ondan sonraki aktif 'ekle' yüklemeleri.
  Aynı ürün + firma + hafta birden fazla geçerli yüklemede varsa en yenisi alınır.
Eski yüklemeler silinmez; geri alınan yükleme hesaptan çıkar, öncekiler tekrar geçerli olur.
aktarim_id'si boş (NULL) satırlar elle girilmiş kabul edilir ve her zaman geçerlidir.
"""
import os
import sqlite3
from datetime import datetime

SCHEMA_VERSION = 3

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
    geri_alma_tarihi TEXT,
    mod           TEXT NOT NULL DEFAULT 'yerine' -- yerine: aynı ayın öncekilerin yerine geçer
        CHECK (mod IN ('yerine', 'ekle')),       -- ekle  : aynı ayın üstüne eklenir (başka müşteri dosyası)
    kullanici     TEXT,                          -- yükleyen: Windows kullanıcısı @ bilgisayar
    geri_alan     TEXT                           -- geri alan kullanıcı
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
CREATE INDEX IF NOT EXISTS ix_siparisler_anahtar ON siparisler(urun_kodu, firma, yil, hafta);

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
CREATE INDEX IF NOT EXISTS ix_backlog_anahtar ON backlog(urun_kodu, firma, yil, ay);

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

"""

# Görünümler her açılışta yeniden kurulur (veri tutmazlar).
SCHEMA_VIEWS = """
DROP VIEW IF EXISTS gecerli_backlog;
DROP VIEW IF EXISTS gecerli_siparisler;
DROP VIEW IF EXISTS gecerli_aktarimlar;

-- Hesapta kullanılan yüklemeler (her yıl + ay için):
--   en son aktif 'yerine' yükleme + ondan SONRA gelen aktif 'ekle' yüklemeleri
CREATE VIEW gecerli_aktarimlar AS
SELECT a.* FROM aktarimlar a
WHERE a.durum = 'aktif'
  AND a.id >= COALESCE((SELECT MAX(b.id) FROM aktarimlar b
                        WHERE b.durum = 'aktif' AND b.yil = a.yil AND b.ay = a.ay
                          AND b.mod = 'yerine'), 0);

-- Aynı ürün + firma + hafta birden fazla geçerli yüklemede varsa EN YENİSİ alınır (çift sayım olmaz)
CREATE VIEW gecerli_siparisler AS
SELECT s.* FROM siparisler s
WHERE s.aktarim_id IS NULL
   OR s.aktarim_id = (SELECT MAX(s2.aktarim_id) FROM siparisler s2
                      WHERE s2.aktarim_id IN (SELECT id FROM gecerli_aktarimlar)
                        AND s2.urun_kodu = s.urun_kodu AND s2.firma = s.firma
                        AND s2.yil = s.yil AND s2.hafta = s.hafta);

CREATE VIEW gecerli_backlog AS
SELECT k.* FROM backlog k
WHERE k.aktarim_id IS NULL
   OR k.aktarim_id = (SELECT MAX(k2.aktarim_id) FROM backlog k2
                      WHERE k2.aktarim_id IN (SELECT id FROM gecerli_aktarimlar)
                        AND k2.urun_kodu = k.urun_kodu AND k2.firma = k.firma
                        AND k2.yil = k.yil AND k2.ay = k.ay);
"""


def current_user() -> str:
    """İşlem kaydı için: Windows kullanıcı adı @ bilgisayar adı (ör. MONSTER@MONSTERPC)."""
    import getpass
    import socket
    try:
        user = getpass.getuser()
    except Exception:
        user = os.environ.get("USERNAME") or os.environ.get("USER") or "?"
    try:
        host = socket.gethostname()
    except Exception:
        host = os.environ.get("COMPUTERNAME") or "?"
    return f"{user}@{host}"


def normalize_firma(value) -> str:
    """Firma adını tek tip yapar: Türkçe büyük harf + tek boşluk. 'Valeo  bursa ' -> 'VALEO BURSA'."""
    if value is None:
        return ""
    s = str(value)
    if s.strip().lower() in ("nan", "none"):
        return ""
    s = " ".join(s.split())
    return s.replace("i", "İ").replace("ı", "I").upper()


def _normalize_firms_in_db(con):
    """Eski kayıtlarda farklı yazılmış firma adlarını tek tipe çevirir (aynı yüklemede çakışırsa adetleri birleştirir)."""
    for table, key_cols, sum_cols in (("siparisler", ("aktarim_id", "urun_kodu", "yil", "hafta"), ("adet",)),
                                      ("backlog", ("aktarim_id", "urun_kodu", "yil", "ay"), ("backlog_adet", "stok_adet"))):
        rows = con.execute(f"SELECT id, firma FROM {table}").fetchall()
        bad = [(rid, f, normalize_firma(f)) for rid, f in rows if (f or "") != normalize_firma(f)]
        for rid, old, new in bad:
            try:
                con.execute(f"UPDATE {table} SET firma=? WHERE id=?", (new, rid))
            except sqlite3.IntegrityError:
                # aynı yüklemede normalize hali zaten var -> adetleri ona ekle, bu satırı kaldır
                r = con.execute(f"SELECT {', '.join(key_cols)}, {', '.join(sum_cols)} FROM {table} WHERE id=?",
                                (rid,)).fetchone()
                keys, vals = r[:len(key_cols)], r[len(key_cols):]
                where = " AND ".join(f"{c} IS ?" for c in key_cols) + " AND firma=?"
                sets = ", ".join(f"{c} = {c} + ?" for c in sum_cols)
                con.execute(f"UPDATE {table} SET {sets} WHERE {where}", (*vals, *keys, new))
                con.execute(f"DELETE FROM {table} WHERE id=?", (rid,))


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
    # timeout: dosya kısa süre kilitliyse (ör. OneDrive senkronu) hemen hata vermek yerine 15 sn bekle
    con = sqlite3.connect(path, timeout=15)
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
        # önceki sürümde 'mod' yoktu: eski yüklemeler 'yerine' sayılır
        if "mod" not in _columns(con, "aktarimlar"):
            con.execute("ALTER TABLE aktarimlar ADD COLUMN mod TEXT NOT NULL DEFAULT 'yerine'")
        for col in ("kullanici", "geri_alan"):     # önceki sürümde kim yaptı bilgisi yoktu
            if col not in _columns(con, "aktarimlar"):
                con.execute(f"ALTER TABLE aktarimlar ADD COLUMN {col} TEXT")
        con.executescript(SCHEMA_VIEWS)
        _normalize_firms_in_db(con)
        con.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        con.commit()
    finally:
        con.close()
    return path


def check_db(path: str):
    """Dosya geçerli bir SQLite veritabanı mı? Sorun varsa açıklama döner, yoksa None.
    (Dosya yoksa None döner – o durum ayrıca ele alınır.)"""
    if not os.path.exists(path):
        return None
    if os.path.isdir(path):
        return "Seçilen yol bir klasör, veritabanı dosyası değil."
    try:
        size = os.path.getsize(path)
        if size:
            with open(path, "rb") as f:
                head = f.read(16)
            if head != b"SQLite format 3\x00":
                return ("Bu dosya geçerli bir veritabanı değil (SQLite dosyası değil). Yanlış dosya seçilmiş "
                        "veya dosya bozulmuş olabilir. Yedekten geri yükleyin ya da doğru .db dosyasını seçin.")
        con = sqlite3.connect(path, timeout=15)
        try:
            con.execute("SELECT count(*) FROM sqlite_master").fetchone()
        finally:
            con.close()
    except sqlite3.DatabaseError as e:
        return f"Veritabanı açılamadı ({e}). Dosya bozulmuş olabilir; yedekten geri yükleyin."
    except OSError as e:
        return f"Dosya okunamadı: {e}"
    return None


def is_network_path(path: str) -> bool:
    """Veritabanı ağ klasöründe mi? (\\\\sunucu\\paylaşım veya ağ sürücüsü)"""
    raw = str(path).strip()
    if raw.startswith("\\\\") or raw.startswith("//"):     # \\sunucu\paylaşım (UNC) – yazıldığı haliyle kontrol
        return True
    p = os.path.abspath(path)
    if p.startswith("\\\\") or p.startswith("//"):
        return True
    if os.name == "nt":
        try:
            import ctypes
            root = os.path.splitdrive(p)[0] + "\\"
            return ctypes.windll.kernel32.GetDriveTypeW(root) == 4   # DRIVE_REMOTE
        except Exception:
            return False
    return False


# --- Aynı veritabanını kimler açmış? (yanında küçük bir işaret dosyası) ---
SESSION_MAX_AGE_H = 12


def _sessions_file(path: str) -> str:
    return os.path.abspath(path) + ".kullanimda"


def _read_sessions(path: str) -> list:
    import json
    try:
        with open(_sessions_file(path), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _write_sessions(path: str, sessions: list):
    import json
    f = _sessions_file(path)
    try:
        if sessions:
            with open(f, "w", encoding="utf-8") as fh:
                json.dump(sessions, fh, ensure_ascii=False)
        elif os.path.exists(f):
            os.remove(f)
    except Exception:
        pass   # işaret dosyası yazılamazsa uygulama yine çalışır


def _fresh(sessions: list) -> list:
    now = datetime.now()
    out = []
    for s in sessions:
        try:
            if (now - datetime.fromisoformat(s["zaman"])).total_seconds() < SESSION_MAX_AGE_H * 3600:
                out.append(s)
        except Exception:
            pass
    return out


def _me() -> dict:
    return {"kullanici": current_user(), "pid": os.getpid()}


def other_sessions(path: str) -> list:
    """Bu veritabanını şu an açık tutan DİĞER oturumlar (12 saatten eskiler sayılmaz)."""
    me = _me()
    return [s for s in _fresh(_read_sessions(path))
            if not (s.get("kullanici") == me["kullanici"] and s.get("pid") == me["pid"])]


def register_session(path: str):
    me = _me()
    sessions = [s for s in _fresh(_read_sessions(path))
                if not (s.get("kullanici") == me["kullanici"] and s.get("pid") == me["pid"])]
    sessions.append({**me, "zaman": datetime.now().isoformat(timespec="seconds")})
    _write_sessions(path, sessions)


def unregister_session(path: str):
    me = _me()
    _write_sessions(path, [s for s in _fresh(_read_sessions(path))
                           if not (s.get("kullanici") == me["kullanici"] and s.get("pid") == me["pid"])])


def is_synced_folder(path: str) -> bool:
    """Veritabanı OneDrive / Dropbox / Google Drive gibi senkron klasöründe mi?"""
    p = os.path.abspath(path).replace("\\", "/").lower()
    return any(k in p for k in ("/onedrive", "/dropbox", "/google drive", "/googledrive", "/icloud"))


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


# ----------------------------------------------------------------------
# Siparişler penceresi: hesaptaki siparişleri görme ve elle düzenleme
# ----------------------------------------------------------------------
AY_ADLARI = ["OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
             "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
MANUAL_FILE = "Elle düzenleme"


def load_valid_orders(path: str):
    """Hesapta kullanılan siparişler (adet > 0), kaynağıyla birlikte."""
    import pandas as pd
    create_empty_db(path)
    con = connect(path)
    try:
        df = pd.read_sql_query("""
            SELECT s.urun_kodu, s.firma, s.yil, s.ay, s.hafta, s.hafta_yili, s.adet, s.aktarim_id,
                   COALESCE(a.dosya_adi, 'elle girilmiş') AS kaynak, COALESCE(a.tarih, '') AS tarih
            FROM gecerli_siparisler s LEFT JOIN aktarimlar a ON a.id = s.aktarim_id
            WHERE s.adet > 0""", con)
    finally:
        con.close()
    df["firma"] = df["firma"].fillna("")
    return df


def save_manual_orders(path: str, changes, aciklama: str = "") -> list:
    """Elle yapılan sipariş değişikliklerini yazar. changes: [{urun_kodu, firma, yil, ay, hafta, adet}, ...]
    adet = 0 -> sipariş silinir (hesaptan çıkar). Her yıl + ay için bir 'Elle düzenleme' yüklemesi (mod='ekle')
    oluşturulur; Aktarımlar sekmesinden geri alınabilir. Tek işlem: hata olursa hiçbiri yazılmaz."""
    from collections import defaultdict
    create_empty_db(path)
    groups = defaultdict(list)
    for c in changes:
        groups[(int(c["yil"]), c["ay"])].append(c)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ids = []
    con = connect(path)
    try:
        with con:
            for (yil, ay), rows in sorted(groups.items(), key=lambda kv: (kv[0][0], AY_ADLARI.index(kv[0][1]))):
                ay_no = AY_ADLARI.index(ay) + 1
                cur = con.execute(
                    "INSERT INTO aktarimlar(tarih, dosya_adi, sayfa, yil, ay, ay_no, haftalar, siparis_satir, "
                    "backlog_satir, mod, kullanici) VALUES (?,?,?,?,?,?,?,?,0,'ekle',?)",
                    (now, MANUAL_FILE, aciklama or None, yil, ay, ay_no,
                     ",".join(str(h) for h in sorted({int(r["hafta"]) for r in rows})),
                     sum(1 for r in rows if float(r["adet"]) > 0), current_user()))
                aid = cur.lastrowid
                con.executemany(
                    "INSERT INTO siparisler(aktarim_id, urun_kodu, firma, yil, hafta, hafta_yili, ay, adet) "
                    "VALUES (?,?,?,?,?,?,?,?)",
                    [(aid, r["urun_kodu"], normalize_firma(r.get("firma")), yil, int(r["hafta"]),
                      hafta_yili(yil, ay_no, int(r["hafta"])), ay, float(r["adet"])) for r in rows])
                ids.append(aid)
    finally:
        con.close()
    return ids


def list_imports(path: str):
    """Tüm yüklemeler, en yeni en üstte. 'gecerli' = hesapta kullanılıyor mu."""
    con = connect(path)
    try:
        rows = con.execute("""
            SELECT a.id, a.tarih, a.dosya_adi, a.sayfa, a.yil, a.ay, a.haftalar,
                   a.siparis_satir, a.backlog_satir, a.durum, a.mod, a.kullanici, a.geri_alan, a.geri_alma_tarihi,
                   CASE WHEN g.id IS NULL THEN 0 ELSE 1 END AS gecerli
            FROM aktarimlar a LEFT JOIN gecerli_aktarimlar g ON g.id = a.id
            ORDER BY a.id DESC""").fetchall()
        cols = ["id", "tarih", "dosya_adi", "sayfa", "yil", "ay", "haftalar",
                "siparis_satir", "backlog_satir", "durum", "mod", "kullanici", "geri_alan", "geri_alma_tarihi",
                "gecerli"]
        return [dict(zip(cols, r)) for r in rows]
    finally:
        con.close()


def undo_import(path: str, aktarim_id: int):
    """Yüklemeyi geri alır (veri silinmez, yükleme 'geri alındı' olur; varsa bir önceki yükleme geçerli olur)."""
    con = connect(path)
    try:
        cur = con.execute("UPDATE aktarimlar SET durum='geri alındı', geri_alma_tarihi=?, geri_alan=? "
                          "WHERE id=? AND durum='aktif'",
                          (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), current_user(), aktarim_id))
        con.commit()
        return cur.rowcount == 1
    finally:
        con.close()


if __name__ == "__main__":
    import sys
    p = sys.argv[1] if len(sys.argv) > 1 else "kapasite.db"
    create_empty_db(p)
    print("Oluşturuldu:", os.path.abspath(p), table_counts(p))

"""
tablo.py — Gerisi Bende ortak tablo okuma / yazma katmanı ("Sizden Gelenler" paketleri)

Kaynağı: library/_ortak/tablo.py; paket klasörlerindeki kopyalar CI tarafından bununla aynı tutulur.

- Platform dışa aktarımlarını okur: CSV / TSV / TXT (UTF-8, UTF-16, Windows-1254), XLSX, ZIP ve GZ içindeki CSV/TXT.
  Ayırıcı (virgül, noktalı virgül, sekme) başlık satırından bulunur.
- Sütunları adlarıyla eşleştirir (Türkçe karakter, büyük-küçük harf, noktalama farkı gözetmeden); sütun sırasına güvenmez.
- Sayıları ABD (1,234.56) ve Türkiye (1.234,56) yazımında okur.
- Excel çıktısı ve diğer paketlerin okuyacağı özet (.ozet.json) yazar.
İnternete bağlanmaz.
"""
from __future__ import annotations

import csv
import gzip
import io
import json
import re
import zipfile
from datetime import date, datetime, timezone
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

_TR = str.maketrans("çğıöşüâîûÇĞİÖŞÜÂÎÛ", "cgiosuaiuCGIOSUAIU")
METIN_UZANTI = {".csv", ".tsv", ".txt"}
TABLO_UZANTI = METIN_UZANTI | {".xlsx", ".xlsm", ".zip", ".gz"}


def katla(s) -> str:
    """'Amount (Merchant Currency)' → 'amount merchant currency'; 'Ürün Adı' → 'urun adi'."""
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s or "").translate(_TR).lower()).split())


def metin(x) -> str:
    return str(x if x is not None else "").strip()


# ----------------------------------------------------------------------------- okuma

def _coz(veri: bytes) -> str:
    if veri.startswith((b"\xff\xfe", b"\xfe\xff")):
        return veri.decode("utf-16")
    if len(veri) > 1 and veri[1:2] == b"\x00" and veri[:1] != b"\x00":  # BOM'suz UTF-16 LE
        return veri.decode("utf-16-le")
    for kod in ("utf-8-sig", "cp1254"):
        try:
            return veri.decode(kod)
        except UnicodeDecodeError:
            continue
    return veri.decode("latin-1")


def _metin_tablo(icerik: str) -> list[list]:
    satirlar = icerik.splitlines()
    ilk = next((s for s in satirlar if s.strip()), "")
    ayirici = max(",;\t", key=ilk.count)
    return list(csv.reader(io.StringIO(icerik, newline=""), delimiter=ayirici))   # tırnak içindeki çok satırlı hücreler korunur


def _xlsx_tablo(kaynak) -> list[list]:
    wb = load_workbook(kaynak, data_only=True, read_only=True)
    try:
        return [list(r) for r in wb.active.iter_rows(values_only=True)]
    finally:
        wb.close()


def tablo_oku(yol: Path) -> list[list]:
    """Tek dosyayı satır listesine çevirir. ZIP içinde birden çok tablo varsa satırları art arda ekler (başlıklar tekrarlanır)."""
    son = yol.suffix.lower()
    if son in {".xlsx", ".xlsm"}:
        satirlar = _xlsx_tablo(yol)
    elif son == ".gz":
        satirlar = _metin_tablo(_coz(gzip.decompress(yol.read_bytes())))
    elif son == ".zip":
        satirlar = []
        with zipfile.ZipFile(yol) as z:
            for ad in sorted(z.namelist()):
                alt = Path(ad).suffix.lower()
                if alt in METIN_UZANTI:
                    satirlar += _metin_tablo(_coz(z.read(ad)))
                elif alt in {".xlsx", ".xlsm"}:
                    satirlar += _xlsx_tablo(io.BytesIO(z.read(ad)))
    else:
        satirlar = _metin_tablo(_coz(yol.read_bytes()))
    return [r for r in satirlar if any(c not in (None, "") and str(c).strip() for c in r)]


def dosyalar(yollar) -> list[Path]:
    """Dosya ve klasörleri (alt klasörler dahil) tablo dosyalarına açar; ~$ ile başlayan Excel kilit dosyalarını atlar."""
    sonuc = []
    for y in yollar or []:
        y = Path(y)
        if y.is_dir():
            sonuc += sorted(p for p in y.rglob("*") if p.is_file() and p.suffix.lower() in TABLO_UZANTI and not p.name.startswith("~$"))
        elif y.exists():
            sonuc.append(y)
        else:
            raise ValueError(f"Dosya bulunamadı: {y}")
    return sonuc


def kayitlar(yol: Path, sozluk: dict, zorunlu: tuple[str, ...]) -> list[dict]:
    """sozluk: {alan: (olası başlık adları, …)}. Başlık satırını bulur; ZIP'te tekrarlanan başlık satırlarını atlar.
    Her kayıtta '_dosya' ve '_satir' bulunur."""
    s = tablo_oku(yol)
    adlar = {alan: [katla(a) for a in v] for alan, v in sozluk.items()}
    for bi, r in enumerate(s[:15]):
        b = [katla(c) for c in r]
        es = {}
        for alan, olasi in adlar.items():
            j = next((b.index(a) for a in olasi if a in b and b.index(a) not in es.values()), None)
            if j is not None:
                es[alan] = j
        if all(z in es for z in zorunlu):
            baslik = [katla(c) for c in r]
            return [{a: None for a in sozluk} | {a: (row[j] if j < len(row) else None) for a, j in es.items()} | {"_dosya": yol.name, "_satir": n}
                    for n, row in enumerate(s[bi + 1:], bi + 2) if [katla(c) for c in row] != baslik]
    gerekli = ", ".join(sozluk[z][0] for z in zorunlu)
    raise ValueError(f"{yol.name}: başlık satırı bulunamadı; gerekli sütunlar: {gerekli}")


def hepsini_oku(yollar, sozluk: dict, zorunlu: tuple[str, ...]) -> list[dict]:
    kayit = []
    for y in dosyalar(yollar):
        kayit += kayitlar(y, sozluk, zorunlu)
    return kayit


# ----------------------------------------------------------------------------- değerler

def sayi(x, ondalik: str | None = None) -> Decimal | None:
    """Sayı okur. ondalik='.' (platform raporları: 1234.56, 1,234.56), ',' (Türkçe: 1.234,56) ya da None (tahmin).
    Parantezli (1.00) negatif sayılır; para birimi ve yüzde işaretleri atılır."""
    if x is None or x == "":
        return None
    if isinstance(x, bool):
        return None
    if isinstance(x, (int, float, Decimal)):
        return Decimal(str(x))
    s = str(x).strip()
    eksi = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not re.search(r"\d", s):
        return None
    if ondalik is None:
        if "," in s and "." in s:
            ondalik = "," if s.rfind(",") > s.rfind(".") else "."
        elif "," in s:
            ondalik = "." if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s) else ","
        else:
            ondalik = "," if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s) else "."
    s = s.replace("." if ondalik == "," else ",", "")
    if ondalik == ",":
        s = s.replace(",", ".")
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    return -d if eksi else d


def yuvarla(x: Decimal, hane: str = "0.01") -> Decimal:
    return x.quantize(Decimal(hane), rounding=ROUND_HALF_UP)


_BICIMLER = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y%m%d",
             "%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y", "%d/%m/%Y %H:%M", "%d/%m/%Y",
             "%b %d, %Y %I:%M:%S %p", "%b %d, %Y %H:%M:%S", "%b %d, %Y", "%d %b %Y", "%B %d, %Y")


def zaman(x, abd: bool = False) -> datetime | None:
    """Tarih / zaman okur ve UTC saat dilimsiz (naive) datetime döndürür.
    abd=True: '10/01/2026' AA/GG/YYYY okunur (App Store Connect raporları). Sayılar UNIX zamanı sayılır
    (saniye, milisaniye ya da mikrosaniye; Firebase 'event_timestamp' mikrosaniyedir)."""
    if x is None or x == "":
        return None
    if isinstance(x, datetime):
        return x.astimezone(timezone.utc).replace(tzinfo=None) if x.tzinfo else x
    if isinstance(x, date):
        return datetime(x.year, x.month, x.day)
    s = str(x).strip()
    if re.fullmatch(r"\d{10}(\d{3}(\d{3})?)?(\.\d+)?", s) and not re.fullmatch(r"\d{8}", s):
        n = float(s)
        bolen = 1e6 if n > 1e14 else 1e3 if n > 1e11 else 1
        return datetime.fromtimestamp(n / bolen, tz=timezone.utc).replace(tzinfo=None)
    s = re.sub(r"\s*(UTC|GMT|Z)$", "", s).replace("T", " ") if not re.fullmatch(r"\d{8}", s) else s
    s = re.sub(r"(\d{2}:\d{2}:\d{2})\s*[+-]\d{2}:?\d{2}$", r"\1", s).strip()   # saat dilimi farkı (yalnız saatten sonra)
    s = re.sub(r"(\d{2}:\d{2}:\d{2})\.\d+$", r"\1", s)                           # saniye kesri (01.09.2026'ya dokunmaz)
    if abd:
        for f in ("%m/%d/%Y", "%m/%d/%Y %H:%M:%S"):
            try:
                return datetime.strptime(s, f)
            except ValueError:
                pass
    for f in _BICIMLER:
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def gun(x, abd: bool = False) -> date | None:
    z = zaman(x, abd)
    return z.date() if z else None


def tr_tarih(d) -> str:
    return d.strftime("%d.%m.%Y") if d else ""


# ----------------------------------------------------------------------------- Excel

BASLIK_DOLGU, BASLIK_YAZI = PatternFill("solid", fgColor="1D1D1F"), Font(bold=True, color="FFFFFF")
KIRMIZI, SARI, MAVI, YESIL, GRI, MOR = "FDE2E1", "FFF4CE", "E8F0FE", "E3F4E1", "EDEDED", "EDE7FF"
ONEM_RENK = {"Yüksek": KIRMIZI, "Orta": SARI, "Bilgi": MAVI}
UST = Alignment(vertical="top", wrap_text=True)
PF, YF, TF, ZF, TAMF = "#,##0.00", "0.0%", "DD.MM.YYYY", "DD.MM.YYYY HH:MM", "#,##0"


def yeni_kitap() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def sayfa(wb, ad: str, basliklar, genislik=(), dondur="A2"):
    ws = wb.create_sheet(ad)
    ws.append(list(basliklar))
    for h in ws[1]:
        h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = dondur
    return ws


def satir(ws, degerler, bicim: dict | None = None, dolgu: dict | None = None, sar=()):
    """bicim = {sütun no: Excel sayı biçimi}, dolgu = {sütun no: renk}, sar = metni kaydırılacak sütunlar."""
    ws.append([float(d) if isinstance(d, Decimal) else d for d in degerler])
    n = ws.max_row
    for j, b in (bicim or {}).items():
        ws.cell(n, j).number_format = b
    for j, renk in (dolgu or {}).items():
        if renk:
            ws.cell(n, j).fill = PatternFill("solid", fgColor=renk)
    for j in sar:
        ws.cell(n, j).alignment = UST
    return n


def uyari(onem: str, tur: str, kim, aciklama: str) -> dict:
    return {"onem": onem, "tur": tur, "kim": metin(kim), "aciklama": aciklama}


def uyari_sayfasi(wb, uyarilar: list[dict]) -> None:
    sira = {"Yüksek": 0, "Orta": 1, "Bilgi": 2}
    ws = sayfa(wb, "Uyarılar", ["Önem", "Tür", "Kim / Ne", "Açıklama"], (9, 34, 30, 100))
    for u in sorted(uyarilar, key=lambda u: (sira[u["onem"]], u["tur"], u["kim"])):
        satir(ws, [u["onem"], u["tur"], u["kim"], u["aciklama"]], dolgu={1: ONEM_RENK[u["onem"]]}, sar=(4,))
    if not uyarilar:
        satir(ws, ["Bilgi", "Uyarı yok", "", "Girdi dosyalarında dikkat gerektiren bir durum bulunmadı."])


def kaydet(wb, cikti: Path) -> None:
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


# ----------------------------------------------------------------------------- özet (paketler arası)

def _json(x):
    if isinstance(x, Decimal):
        return float(x)
    if isinstance(x, (date, datetime)):
        return x.isoformat()
    raise TypeError(type(x))


def ozet_yaz(klasor: Path, paket: str, veri: dict) -> Path:
    """Yönetim özeti paketinin okuyacağı <paket>.ozet.json dosyasını yazar.
    veri: {"gunluk": {"YYYY-MM-DD": {metrik: değer}}, ...}"""
    klasor.mkdir(parents=True, exist_ok=True)
    yol = klasor / f"{paket}.ozet.json"
    icerik = {"paket": paket, "surum": 1, "uretildi": datetime.now().replace(microsecond=0).isoformat(), **veri}
    yol.write_text(json.dumps(icerik, ensure_ascii=False, indent=1, default=_json, sort_keys=False), encoding="utf-8")
    return yol

"""
CV Raporlama — Workers / Workless AI Agent
İnsan Kaynakları › İK Uzmanı

Kod bloğu sürümünün dosya okuyucularını kullanır; her CV'yi yapay zekâ ile
yapılandırılmış bir profile dönüştürür ve (verildiyse) iş ilanına uyumunu
gerekçesiyle puanlar.

Gizlilik: e-posta, telefon, LinkedIn ve ad soyad bilgisayarınızda çıkarılır ve
modele gönderilen metinden maskelenir. Veri gönderilmeden önce onayınız alınır.

Kullanım:
    python agent.py                                   # örnek veriyle dener
    python agent.py --girdi ./cvler --ilan ilan.txt --cikti cikti/aday_raporu_ai.xlsx
"""
from __future__ import annotations

import argparse
import sys
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import cv_cekirdek as cekirdek
import llm

BURASI = Path(__file__).resolve().parent
MAKS_KARAKTER = 60_000  # tek bir CV için üst sınır (çok uzun dosyalar uyarıyla kesilir)

SEMA = {
    "type": "object",
    "properties": {
        "son_pozisyon": {"type": "string"},
        "son_sirket": {"type": "string"},
        "toplam_deneyim_yil": {"type": "number"},
        "en_yuksek_egitim": {"type": "string", "enum": ["Doktora", "Yüksek Lisans", "Lisans", "Ön Lisans", "Lise", "Belirsiz"]},
        "bolum": {"type": "string"},
        "diller": {"type": "array", "items": {"type": "string"}},
        "beceriler": {"type": "array", "items": {"type": "string"}},
        "ozet": {"type": "string"},
        "ilan_uyum_puani": {"anyOf": [{"type": "integer"}, {"type": "null"}]},
        "guclu_yonler": {"type": "array", "items": {"type": "string"}},
        "eksikler": {"type": "array", "items": {"type": "string"}},
        "gerekce": {"type": "string"},
    },
    "required": ["son_pozisyon", "son_sirket", "toplam_deneyim_yil", "en_yuksek_egitim", "bolum", "diller",
                 "beceriler", "ozet", "ilan_uyum_puani", "guclu_yonler", "eksikler", "gerekce"],
    "additionalProperties": False,
}


@dataclass
class CV:
    dosya: str
    metin: str = ""
    ad_soyad: str = ""
    eposta: str = ""
    telefon: str = ""
    linkedin: str = ""
    profil: dict = field(default_factory=dict)
    uyarilar: list[str] = field(default_factory=list)


def hazirla(yol: Path, maske: bool = True) -> CV:
    """Dosyayı okur, iletişim bilgisini yerelde çıkarır ve gönderilecek metni maskeler."""
    cv = CV(dosya=yol.name)
    try:
        ham = cekirdek.metin_oku(yol)
    except Exception as hata:
        cv.uyarilar.append(f"Okunamadı: {hata.__class__.__name__}")
        return cv
    if len(ham.strip()) < 40:
        cv.uyarilar.append("Metin çıkarılamadı (taranmış görüntü PDF olabilir, OCR gerekir)")
        return cv

    cv.ad_soyad = cekirdek.ad_bul(ham, yol.name)
    e = cekirdek.EPOSTA.search(ham)
    cv.eposta = e.group(0) if e else ""
    cv.telefon = cekirdek.telefon_bul(ham)
    li = cekirdek.LINKEDIN.search(ham)
    cv.linkedin = li.group(0) if li else ""
    return cv_metni_ata(cv, ham, maske)


def cv_metni_ata(cv: CV, ham: str, maske: bool = True) -> CV:
    metin = ham
    if maske:
        metin = llm.maskele(metin)
        for parca in {cv.ad_soyad, cv.ad_soyad.upper()} - {""}:
            metin = metin.replace(parca, "[AD SOYAD]")
    if len(metin) > MAKS_KARAKTER:
        metin = metin[:MAKS_KARAKTER]
        cv.uyarilar.append(f"CV çok uzun; ilk {MAKS_KARAKTER:_} karakter analiz edildi".replace("_", "."))
    cv.metin = metin
    return cv


def analiz_et(cv: CV, sistem: str, ilan: str | None) -> CV:
    bugun = datetime.now().strftime("%Y-%m-%d")
    parcalar = [f"Bugünün tarihi: {bugun}"]
    parcalar.append(f"<is_ilani>\n{ilan}\n</is_ilani>" if ilan else "İş ilanı verilmedi; ilan_uyum_puani null olmalı.")
    parcalar.append(f"<cv dosya=\"{cv.dosya}\">\n{cv.metin}\n</cv>")
    cv.profil = llm.json_iste(sistem, "\n\n".join(parcalar), SEMA)
    if not ilan:
        cv.profil["ilan_uyum_puani"] = None
    return cv


# ----------------------------------------------------------------------------
# Excel çıktısı
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
KARAR_DOLGU = PatternFill("solid", fgColor="FFF4CE")


def rapor_yaz(cvler: list[CV], cikti: Path, ilan_var: bool) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Adaylar"
    sutunlar = [
        ("Dosya", 22), ("Ad Soyad", 20), ("E-posta", 26), ("Telefon", 17), ("Son Pozisyon", 24), ("Son Şirket", 22),
        ("Deneyim (yıl)", 10), ("Eğitim", 14), ("Bölüm", 20), ("Diller", 18), ("Beceriler", 42), ("Özet", 60),
    ]
    if ilan_var:
        sutunlar += [("İlan Uyum Puanı", 10), ("Güçlü Yönler", 40), ("Eksikler", 40), ("Gerekçe", 60)]
    sutunlar += [("Uyarılar", 30), ("İnsan Kararı", 16)]
    ws.append([s for s, _ in sutunlar])
    for i, (_, g) in enumerate(sutunlar, 1):
        ws.column_dimensions[get_column_letter(i)].width = g
        ws.cell(1, i).fill, ws.cell(1, i).font = BASLIK_DOLGU, BASLIK_YAZI

    def sirala(c: CV):
        return (-(c.profil.get("ilan_uyum_puani") or -1), -(c.profil.get("toplam_deneyim_yil") or 0))

    for c in sorted(cvler, key=sirala):
        p = c.profil
        satir = [c.dosya, c.ad_soyad, c.eposta, c.telefon, p.get("son_pozisyon", ""), p.get("son_sirket", ""),
                 p.get("toplam_deneyim_yil"), p.get("en_yuksek_egitim", ""), p.get("bolum", ""),
                 ", ".join(p.get("diller", [])), ", ".join(p.get("beceriler", [])), p.get("ozet", "")]
        if ilan_var:
            satir += [p.get("ilan_uyum_puani"), "\n".join(p.get("guclu_yonler", [])),
                      "\n".join(p.get("eksikler", [])), p.get("gerekce", "")]
        satir += ["; ".join(c.uyarilar), ""]
        ws.append(satir)
    for satir in ws.iter_rows(min_row=2):
        for h in satir:
            h.alignment = Alignment(vertical="top", wrap_text=True)
        satir[-1].fill = KARAR_DOLGU
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    notlar = wb.create_sheet("Notlar")
    for s in [
        ["Oluşturulma", datetime.now().strftime("%d.%m.%Y %H:%M")],
        ["Model", llm.kullanim_ozeti()],
        ["Önemli", "Bu rapor yapay zekâ tarafından üretilmiş bir karar desteğidir. Puanlar hatalı olabilir."],
        ["", "Hiçbir aday yalnızca bu rapora dayanarak elenmemelidir. 'İnsan Kararı' sütununu siz doldurun."],
        ["Gizlilik", "İletişim bilgileri ve ad soyad modele gönderilmeden önce maskelenmiştir."],
    ]:
        notlar.append(s)
    notlar.column_dimensions["A"].width = 14
    notlar.column_dimensions["B"].width = 100

    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


# ----------------------------------------------------------------------------
# Komut satırı
# ----------------------------------------------------------------------------

def calistir(girdi: Path, cikti: Path, ilan_yolu: Path | None, evet: bool = False, maske: bool = True, paralel: int = 4) -> list[CV]:
    if not girdi.is_dir():
        raise llm.LLMHatasi(f"Girdi klasörü bulunamadı: {girdi}")
    dosyalar = sorted(p for p in girdi.iterdir() if p.suffix.lower() in cekirdek.DESTEKLENEN)
    if not dosyalar:
        raise llm.LLMHatasi(f"{girdi} içinde PDF, DOCX veya TXT dosyası yok.")

    cvler = [hazirla(p, maske) for p in dosyalar]
    gonderilecek = [c for c in cvler if c.metin]
    ilan = cekirdek.txt_oku(ilan_yolu).strip() if ilan_yolu else None

    if gonderilecek:
        toplam = sum(len(c.metin) for c in gonderilecek) + (len(ilan) if ilan else 0) * len(gonderilecek)
        gizlilik = "iletişim bilgileri ve ad soyad maskelenmiş" if maske else "MASKELEME KAPALI"
        adet = f"{toplam:_}".replace("_", ".")
        llm.onay_al(f"{len(gonderilecek)} CV metni ({gizlilik}, ~{adet} karakter) analiz için gönderilecek.", evet)

        sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
        # ilk CV'yi tek başına çalıştır: anahtar/model hatası varsa hemen dur
        analiz_et(gonderilecek[0], sistem, ilan)
        print(f"  [1/{len(gonderilecek)}] {gonderilecek[0].dosya}")

        def tek(c: CV) -> CV:
            try:
                analiz_et(c, sistem, ilan)
            except llm.LLMHatasi as h:
                c.uyarilar.append(f"Analiz edilemedi: {h}")
            return c

        with ThreadPoolExecutor(max_workers=max(1, paralel)) as havuz:
            for i, c in enumerate(havuz.map(tek, gonderilecek[1:]), start=2):
                print(f"  [{i}/{len(gonderilecek)}] {c.dosya}")

    rapor_yaz(cvler, cikti, ilan_var=bool(ilan))
    return cvler


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()  # çalışılan klasördeki .env (varsa) — mevcut değerleri ezmez
    except ImportError:
        pass

    p = argparse.ArgumentParser(description="CV'leri yapay zekâ ile analiz edip Excel raporu üretir.")
    p.add_argument("--girdi", type=Path, default=BURASI / "ornek_veri" / "cvler", help="CV klasörü")
    p.add_argument("--ilan", type=Path, default=None, help="İş ilanı metni (.txt)")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "aday_raporu_ai.xlsx", help="Excel çıktı yolu")
    p.add_argument("--paralel", type=int, default=4, help="Aynı anda kaç CV analiz edilsin")
    p.add_argument("--maskeleme-kapali", action="store_true", help="Kişisel bilgileri maskelemeden gönder (önerilmez)")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)

    if a.ilan is None and a.girdi == BURASI / "ornek_veri" / "cvler":
        a.ilan = BURASI / "ornek_veri" / "ilan.txt"   # örnek çalıştırmada örnek ilanı da kullan
    try:
        cvler = calistir(a.girdi, a.cikti, a.ilan, a.evet, maske=not a.maskeleme_kapali, paralel=a.paralel)
    except llm.LLMHatasi as h:
        print(f"[X] {h}")
        return 1
    print(f"\n[OK] {len(cvler)} CV işlendi ({sum(1 for c in cvler if c.uyarilar)} tanesinde uyarı var)")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

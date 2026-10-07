"""
Banka Mutabakatı — Workers / Workless AI Agent
Finans & Muhasebe › Muhasebe Elemanı

1. Eşleştirmeyi kod bloğunun deterministik motoru yapar (yapay zekâ yok).
2. Yalnızca eşleşmeyen açık kalemler modele gönderilir; model her biri için olası
   nedeni, ilişkili kalemi ve düzeltme yevmiye kaydı taslağını yazar.
3. Sonuç kod bloğunun Excel raporuna eklenir; her öneri için 'İnsan Onayı' sütunu vardır.

Kullanım:
    python agent.py                                    # örnek veriyle dener
    python agent.py --banka ekstre.xlsx --defter muavin.xlsx --cikti cikti/mutabakat_ai.xlsx
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import llm
import mutabakat_cekirdek as cekirdek

BURASI = Path(__file__).resolve().parent
PAKET = 60  # bir istekte gönderilen en fazla açık kalem

NEDENLER = ["Banka masrafı/komisyon", "Faiz/getiri", "Zamanlama farkı", "Tutar hatası", "Mükerrer kayıt",
            "Bankaya yansımamış işlem", "Deftere işlenmemiş işlem", "Diğer"]

SEMA = {
    "type": "object",
    "properties": {
        "kalemler": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "etiket": {"type": "string"},
                    "olasi_neden": {"type": "string", "enum": NEDENLER},
                    "iliskili_kalem": {"type": "string"},
                    "aciklama": {"type": "string"},
                    "onerilen_kayit": {"type": "string"},
                    "guven": {"type": "string", "enum": ["dusuk", "orta", "yuksek"]},
                },
                "required": ["etiket", "olasi_neden", "iliskili_kalem", "aciklama", "onerilen_kayit", "guven"],
                "additionalProperties": False,
            },
        },
        "genel_degerlendirme": {"type": "string"},
    },
    "required": ["kalemler", "genel_degerlendirme"],
    "additionalProperties": False,
}


def kalem_satiri(k: cekirdek.Kayit, karsi, ayni, tol: int) -> str:
    on_tahmin = cekirdek.olasi_neden(k, karsi, ayni, tol)
    return f"{k.etiket} | {k.tarih:%d.%m.%Y} | {llm.maskele(k.aciklama)} | {cekirdek.tl(k.tutar)} | ön tahmin: {on_tahmin}"


def degerlendir(banka, defter, acik_b, acik_d, tol: int) -> tuple[dict[str, dict], list[str]]:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    baglam = (f"Banka hareket sayısı: {len(banka)}, defter kayıt sayısı: {len(defter)}, "
              f"fark (banka − defter): {cekirdek.tl(sum(k.tutar for k in banka) - sum(k.tutar for k in defter))}, "
              f"tarih toleransı: {tol} gün.")
    tum = [("banka", k) for k in acik_b] + [("defter", k) for k in acik_d]
    sonuc, yorumlar = {}, []
    for i in range(0, len(tum), PAKET):
        parca = tum[i:i + PAKET]
        satirlar_b = [kalem_satiri(k, acik_d, banka, tol) for t, k in parca if t == "banka"]
        satirlar_d = [kalem_satiri(k, acik_b, defter, tol) for t, k in parca if t == "defter"]
        # karşı taraftaki tüm açık kalemleri bağlam olarak ver (ilişki kurabilmesi için)
        bag_b = [kalem_satiri(k, acik_d, banka, tol) for k in acik_b if ("banka", k) not in parca]
        bag_d = [kalem_satiri(k, acik_b, defter, tol) for k in acik_d if ("defter", k) not in parca]
        mesaj = "\n".join([
            baglam,
            "\n<degerlendirilecek_banka_kalemleri>", *satirlar_b, "</degerlendirilecek_banka_kalemleri>",
            "\n<degerlendirilecek_defter_kalemleri>", *satirlar_d, "</degerlendirilecek_defter_kalemleri>",
            "\n<baglam_icin_diger_acik_kalemler>", *(bag_b + bag_d), "</baglam_icin_diger_acik_kalemler>",
            "\nYalnızca 'degerlendirilecek' bölümlerindeki kalemler için sonuç döndür.",
        ])
        yanit = llm.json_iste(sistem, mesaj, SEMA)
        gecerli = {k.etiket for _, k in parca}
        for s in yanit.get("kalemler", []):
            if s.get("etiket") in gecerli:  # modelin uydurduğu numaraları yok say
                sonuc[s["etiket"]] = s
        if yanit.get("genel_degerlendirme"):
            yorumlar.append(yanit["genel_degerlendirme"])
    return sonuc, yorumlar


GUVEN = {"dusuk": "Düşük", "orta": "Orta", "yuksek": "Yüksek"}
AI_DOLGU = PatternFill("solid", fgColor="EDE7FF")
ONAY_DOLGU = PatternFill("solid", fgColor="FFF4CE")
BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")


def rapora_ekle(cikti: Path, sonuc: dict[str, dict], yorumlar: list[str]) -> None:
    wb = load_workbook(cikti)
    ek = [("AI Nedeni", 24), ("İlişkili Kalem", 12), ("AI Açıklama", 60), ("Önerilen Kayıt (taslak)", 56), ("Güven", 10), ("İnsan Onayı", 14)]
    for ad in ("Açık - Banka", "Açık - Defter"):
        ws = wb[ad]
        bas = ws.max_column + 1
        for j, (baslik, gen) in enumerate(ek):
            h = ws.cell(1, bas + j, baslik)
            h.fill, h.font = BASLIK_DOLGU, BASLIK_YAZI
            ws.column_dimensions[h.column_letter].width = gen
        for r in range(2, ws.max_row + 1):
            s = sonuc.get(ws.cell(r, 1).value)
            degerler = ([s["olasi_neden"], s["iliskili_kalem"], s["aciklama"], s["onerilen_kayit"], GUVEN.get(s["guven"], s["guven"]), ""]
                        if s else ["AI yanıt vermedi", "", "", "", "", ""])
            for j, v in enumerate(degerler):
                h = ws.cell(r, bas + j, v)
                h.alignment = Alignment(vertical="top", wrap_text=True)
                h.fill = ONAY_DOLGU if j == len(ek) - 1 else AI_DOLGU
        ws.auto_filter.ref = ws.dimensions

    notlar = wb.create_sheet("AI Değerlendirme", 1)
    satirlar = [["Genel değerlendirme", y] for y in yorumlar] + [
        ["Model", llm.kullanim_ozeti()],
        ["Oluşturulma", datetime.now().strftime("%d.%m.%Y %H:%M")],
        ["Önemli", "Eşleştirme deterministiktir; mor sütunlar yapay zekâ önerisidir. Yevmiye kayıtları taslaktır, "
                   "muhasebeleştirmeden önce kontrol edip 'İnsan Onayı' sütununu doldurun."],
    ]
    for s in satirlar:
        notlar.append(s)
    notlar.column_dimensions["A"].width = 22
    notlar.column_dimensions["B"].width = 110
    for satir in notlar.iter_rows():
        for h in satir:
            h.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(cikti)


def calistir(banka_yolu: Path, defter_yolu: Path, cikti: Path, gun_toleransi: int = 3, evet: bool = False):
    banka = cekirdek.dosya_oku(banka_yolu, "banka")
    defter = cekirdek.dosya_oku(defter_yolu, "defter")
    eslesmeler, acik_b, acik_d = cekirdek.eslestir(banka, defter, gun_toleransi)
    cekirdek.rapor_yaz(cikti, banka, defter, eslesmeler, acik_b, acik_d, gun_toleransi)
    print(f"[OK] Deterministik eşleştirme: {len(eslesmeler)} eşleşti · açık kalem: banka {len(acik_b)}, defter {len(acik_d)}")

    sonuc, yorumlar = {}, []
    if acik_b or acik_d:
        llm.onay_al(f"Yalnızca {len(acik_b) + len(acik_d)} açık kalemin tarih, açıklama (IBAN/TCKN maskeli) "
                    "ve tutarı yorum için gönderilecek. Eşleşen kayıtlar gönderilmez.", evet)
        sonuc, yorumlar = degerlendir(banka, defter, acik_b, acik_d, gun_toleransi)
    else:
        yorumlar = ["Tüm kayıtlar eşleşti; açık kalem yok."]
    rapora_ekle(cikti, sonuc, yorumlar)
    return eslesmeler, acik_b, acik_d, sonuc


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass

    p = argparse.ArgumentParser(description="Banka mutabakatı + açık kalemler için yapay zekâ yorumu.")
    p.add_argument("--banka", type=Path, default=BURASI / "ornek_veri" / "banka_ekstresi.csv")
    p.add_argument("--defter", type=Path, default=BURASI / "ornek_veri" / "defter_102.csv")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "mutabakat_ai.xlsx")
    p.add_argument("--gun-toleransi", type=int, default=3)
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    try:
        _, acik_b, acik_d, sonuc = calistir(a.banka, a.defter, a.cikti, a.gun_toleransi, a.evet)
    except llm.LLMHatasi as h:
        print(f"[X] {h}")
        return 1
    print(f"[OK] {len(sonuc)}/{len(acik_b) + len(acik_d)} açık kalem yorumlandı")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    if sonuc:
        print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

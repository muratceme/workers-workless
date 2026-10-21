"""
Uygulama Destek E-postası Sınıflandırma — Gerisi Bende AI Agent (Sizden Gelenler)

1. Kod e-postaları okur (CSV / Excel dökümü, .mbox ya da .eml klasörü), alıntılanan eski yazışmayı ayıklar ve
   hassas konuları kendisi bulur: iade / geri ödeme, hesap ya da veri silme (KVKK), yasal süreç tehdidi,
   "ödeme alındı ama abonelik açılmadı", hesap güvenliği. Gönderen adresi modele gitmez; e-posta, telefon, IBAN,
   TCKN maskelenir.
2. Model her e-postayı kategori, duygu ve aciliyete ayırır, kısa özet ve müşteriye cevap taslağı yazar.
3. Kod kuralları modelin üstündedir:
     - Hassas konu bulunan e-posta "İnsan onayı gerekli" olur; model bunu kaldıramaz. İade bulunan e-postanın
       kategorisi model ne derse desin "İade talebi" olur.
     - Öncelik kodla: P1 = ödeme alındı / abonelik açılmadı, hesap güvenliği, yasal tehdit ya da yüksek aciliyetli
       ödeme / giriş sorunu; P2 = diğer hassas konular, hata / çökme, yüksek aciliyet; P3 = diğerleri.
     - Onay gerektiren e-postanın taslağında söz verme ("iade ettik", "hesabınız silindi"…) varsa taslak işaretlenir.
     - İlk yanıt hedefi SLA saatlerinden (P1 4, P2 24, P3 72 saat; --sla ile değişir).
Kullanım:
    python agent.py                                      # örnek e-postalarla dener
    python agent.py --girdi destek_epostalari.csv
    python agent.py --girdi Gmail.mbox --sla sla.json
"""
from __future__ import annotations

import argparse
import email
import json
import mailbox
import re
import sys
from collections import Counter
from datetime import datetime, timedelta
from email.policy import default as POLICY
from email.utils import parsedate_to_datetime
from pathlib import Path

import llm
import tablo as T

BURASI = Path(__file__).resolve().parent
ORNEK = BURASI / "ornek_veri"
PAKET_ADI = "uygulama-destek-epostasi"
PAKET = 15
KATEGORILER = ["Giriş / hesap", "Abonelik / ödeme", "İade talebi", "Hesap / veri silme (KVKK)", "Hata / çökme", "Özellik isteği",
               "Kullanım sorusu", "İçerik / yapay zekâ cevabı", "Geri bildirim / teşekkür", "Spam / alakasız"]
HASSAS_KATEGORI = {"İade talebi", "Hesap / veri silme (KVKK)"}
SEVIYE = ["dusuk", "orta", "yuksek"]
DUYGU = ["olumlu", "notr", "olumsuz", "ofkeli"]
VARSAYILAN_SLA = {"P1": 4, "P2": 24, "P3": 72}

KURAL = {
    "iade": re.compile(r"(\biade\b|iadesi|iadeyi|param[ıi]? geri|paramı iste|geri ödeme|geri öde|refund|chargeback|ters ibraz|itiraz edeceğim)", re.I),
    "silme": re.compile(r"(hesab[ıi]m[ıi]? (sil|kapat)|hesab[ıi]m[ıi]n silin|verilerim[ıi]? sil|kişisel veri|\bkvkk\b|\bgdpr\b|unutulma hakk|"
                        r"delete my (account|data)|verilerimin silin|üyeliğimi sil)", re.I),
    "yasal": re.compile(r"(avukat|dava (açaca|açacağım)|tüketici hakem|savcılı|\bbtk\b|\bcimer\b|şikayet edeceğim|şikâyet edeceğim|yasal yol)", re.I),
    "odeme_acilmadi": re.compile(r"(ücret|para|ödeme|tutar|çekim)[^.\n]{0,60}(çekildi|alındı|kesildi|ödedim|yansıdı)[^\n]{0,80}"
                                 r"(açılmadı|aktif değil|aktifleşmedi|gelmedi|görünmüyor|tanımlanmadı|kullanamıyorum)", re.I),
    "guvenlik": re.compile(r"(hesab[ıi]m (çalındı|ele geçirildi|hacklendi)|şifrem (değişmiş|değiştirilmiş)|tanımadığım (bir )?(cihaz|giriş)|izinsiz giriş)", re.I),
}
KURAL_AD = {"iade": "iade / geri ödeme", "silme": "hesap / veri silme", "yasal": "yasal süreç", "odeme_acilmadi": "ödeme alındı, abonelik açılmadı",
            "guvenlik": "hesap güvenliği"}
SOZ = re.compile(r"(iade (ettik|edildi|edeceğiz|edilecek|yaptık|yapıldı|yapılacak)|paranız[ıi]? (iade|geri)|ücretiniz[ıi]? (iade|geri)|"
                 r"hesabınız[ıi]? (sildik|silindi|silinecek|kapatıldı)|verileriniz[ıi]? (sildik|silindi|silinecek)|"
                 r"ücretsiz (tanımladık|tanımlanacak|veriyoruz)|(bir )?aylık hediye)", re.I)
ALINTI = re.compile(r"^(On .{5,120} wrote:|.{5,120} tarihinde .{0,120}(şunu yazdı|yazdı):|-{2,}\s*Original Message|-{2,}\s*Orijinal ileti|"
                    r"From: .+|Kimden: .+|Gönderen: .+)\s*$", re.I | re.M)


def govde_temizle(metin: str) -> str:
    """Alıntılanan eski yazışmayı ve '>' ile başlayan satırları atar."""
    m = ALINTI.search(metin)
    if m:
        metin = metin[:m.start()]
    return "\n".join(s for s in metin.splitlines() if not s.lstrip().startswith(">")).strip()


def _eposta_metin(msg) -> str:
    try:
        bolum = msg.get_body(preferencelist=("plain", "html"))
        if bolum is None:
            return ""
        icerik = bolum.get_content()
        if bolum.get_content_type() == "text/html":
            icerik = re.sub(r"<br\s*/?>|</p>", "\n", icerik, flags=re.I)
            icerik = re.sub(r"<[^>]+>", " ", icerik)
        return icerik
    except (KeyError, LookupError):
        return ""


def _eposta_kayit(msg, no: str) -> dict:
    try:
        z = parsedate_to_datetime(msg["date"]) if msg["date"] else None
        z = z.replace(tzinfo=None) if z else None          # gönderenin yerel saati
    except (TypeError, ValueError):
        z = None
    return {"id": no, "zaman": z, "gonderen": str(msg["from"] or ""), "konu": str(msg["subject"] or ""), "mesaj": govde_temizle(_eposta_metin(msg))}


EPOSTA = {"id": ("Talep No", "Kayıt No", "ID", "No", "Message-ID"), "zaman": ("Tarih", "Zaman", "Date", "Gönderim Tarihi"),
          "gonderen": ("Gönderen", "From", "E-posta"), "konu": ("Konu", "Subject"), "mesaj": ("Mesaj", "İçerik", "Body", "Metin")}


def epostalari_oku(yollar) -> list[dict]:
    kayit = []
    for y in [Path(p) for p in yollar]:
        if y.is_dir():
            for i, f in enumerate(sorted(y.glob("*.eml")), 1):
                kayit.append(_eposta_kayit(email.message_from_bytes(f.read_bytes(), policy=POLICY), f.stem))
            continue
        if not y.exists():
            raise ValueError(f"Dosya bulunamadı: {y}")
        if y.suffix.lower() == ".mbox":
            for i, msg in enumerate(mailbox.mbox(y, factory=lambda f: email.message_from_binary_file(f, policy=POLICY)), 1):
                kayit.append(_eposta_kayit(msg, f"M{i}"))
        elif y.suffix.lower() == ".eml":
            kayit.append(_eposta_kayit(email.message_from_bytes(y.read_bytes(), policy=POLICY), y.stem))
        else:
            for i, r in enumerate(T.kayitlar(y, EPOSTA, ("mesaj",)), 1):
                kayit.append({"id": T.metin(r["id"]) or f"E{i}", "zaman": T.zaman(r["zaman"]), "gonderen": T.metin(r["gonderen"]),
                              "konu": T.metin(r["konu"]), "mesaj": govde_temizle(T.metin(r["mesaj"]))})
    gorulen = Counter()
    for k in kayit:                                   # kimlikler tekil olsun
        gorulen[k["id"]] += 1
        if gorulen[k["id"]] > 1:
            k["id"] = f"{k['id']}-{gorulen[k['id']]}"
    return [k for k in kayit if k["mesaj"] or k["konu"]]


def kurallar(e: dict) -> list[str]:
    metin = f"{e['konu']}\n{e['mesaj']}".replace("İ", "i").lower()      # Türkçe büyük İ, re.I ile eşleşmez
    return [ad for ad, d in KURAL.items() if d.search(metin)]


SEMA = {
    "type": "object",
    "properties": {"epostalar": {"type": "array", "items": {
        "type": "object",
        "properties": {
            "id": {"type": "string"},
            "kategori": {"type": "string", "enum": KATEGORILER},
            "duygu": {"type": "string", "enum": DUYGU},
            "aciliyet": {"type": "string", "enum": SEVIYE},
            "ozet": {"type": "string"},
            "cevap_taslagi": {"type": "string"},
        },
        "required": ["id", "kategori", "duygu", "aciliyet", "ozet", "cevap_taslagi"],
        "additionalProperties": False}}},
    "required": ["epostalar"],
    "additionalProperties": False,
}


def siniflandir(epostalar: list[dict], urun: str) -> dict[str, dict]:
    sistem = (BURASI / "prompt.md").read_text(encoding="utf-8")
    sonuc = {}
    for i in range(0, len(epostalar), PAKET):
        parca = epostalar[i:i + PAKET]
        satir = []
        for e in parca:
            ipucu = "; ".join(KURAL_AD[k] for k in e["kural"]) or "-"
            satir.append(f'<eposta id="{e["id"]}" ipuclari="{ipucu}">\nKonu: {llm.maskele(e["konu"])}\n{llm.maskele(e["mesaj"][:4000])}\n</eposta>')
        mesaj = "\n".join([f"<urun>{urun}</urun>", "<kategoriler>" + "; ".join(KATEGORILER) + "</kategoriler>", "<epostalar>", *satir, "</epostalar>"])
        yanit = llm.json_iste(sistem, mesaj, SEMA)
        gecerli = {e["id"] for e in parca}
        for s in yanit.get("epostalar", []):
            if s.get("id") in gecerli:
                sonuc[s["id"]] = s
    return sonuc


def kural_uygula(e: dict, s: dict, sla: dict) -> list[str]:
    duz = []
    k = set(e["kural"])
    if "iade" in k and s["kategori"] != "İade talebi":
        duz.append(f"kategori {s['kategori']} → İade talebi (kural)")
        s["kategori"] = "İade talebi"
    if "silme" in k and "iade" not in k and s["kategori"] != "Hesap / veri silme (KVKK)":
        duz.append(f"kategori {s['kategori']} → Hesap / veri silme (KVKK) (kural)")
        s["kategori"] = "Hesap / veri silme (KVKK)"
    s["onay"] = bool(k & {"iade", "silme", "yasal", "guvenlik"}) or s["kategori"] in HASSAS_KATEGORI
    if k & {"odeme_acilmadi", "guvenlik", "yasal"} or (s["aciliyet"] == "yuksek" and s["kategori"] in ("Abonelik / ödeme", "Giriş / hesap")):
        p = "P1"
    elif s["onay"] or s["kategori"] == "Hata / çökme" or s["aciliyet"] == "yuksek":
        p = "P2"
    else:
        p = "P3"
    s["oncelik"] = p
    s["hedef"] = e["zaman"] + timedelta(hours=sla.get(p, VARSAYILAN_SLA[p])) if e["zaman"] else None
    s["soz"] = ""
    if s["onay"]:
        m = SOZ.search(s["cevap_taslagi"])
        if m:
            s["soz"] = m.group(0)
            duz.append(f"taslakta söz var: \"{m.group(0)}\" — göndermeden önce düzeltin")
    return duz


AI_DOLGU = T.MOR
P_DOLGU = {"P1": "F8C4C1", "P2": T.KIRMIZI, "P3": T.GRI}
DUYGU_AD = {"olumlu": "Olumlu", "notr": "Nötr", "olumsuz": "Olumsuz", "ofkeli": "Öfkeli"}


def rapor_yaz(cikti: Path, epostalar, sonuc, duzeltmeler) -> None:
    wb = T.yeni_kitap()
    bas = ["No", "Zaman", "Gönderen", "Konu", "Mesaj", "Öncelik", "İnsan Onayı", "Kategori", "Duygu", "Kod Bulguları", "Özet", "Cevap Taslağı",
           "İlk Yanıt Hedefi", "Kural Düzeltmesi", "Onay"]
    gen = (9, 16, 24, 26, 60, 8, 11, 22, 9, 26, 40, 70, 16, 40, 8)
    sirali = sorted(epostalar, key=lambda e: (sonuc.get(e["id"], {}).get("oncelik", "P9"), e["zaman"] or datetime.min))
    for ad, filtre in (("Onay Bekleyenler", lambda s: s and s["onay"]), ("Tüm E-postalar", lambda s: True)):
        ws = T.sayfa(wb, ad, bas, gen, "F2")
        for e in sirali:
            s = sonuc.get(e["id"])
            if not filtre(s):
                continue
            if not s:
                T.satir(ws, [e["id"], e["zaman"], e["gonderen"], e["konu"], e["mesaj"], "", "", "", "", "; ".join(KURAL_AD[k] for k in e["kural"]), "AI yanıt vermedi"],
                        bicim={2: T.ZF}, sar=(5,))
                continue
            T.satir(ws, [e["id"], e["zaman"], e["gonderen"], e["konu"], e["mesaj"], s["oncelik"], "GEREKLİ" if s["onay"] else "", s["kategori"], DUYGU_AD[s["duygu"]],
                         "; ".join(KURAL_AD[k] for k in e["kural"]), s["ozet"], s["cevap_taslagi"], s["hedef"], "; ".join(duzeltmeler.get(e["id"], [])), ""],
                    bicim={2: T.ZF, 13: T.ZF}, sar=(5, 11, 12, 14),
                    dolgu={6: P_DOLGU[s["oncelik"]], 7: T.SARI if s["onay"] else None, 8: AI_DOLGU, 9: AI_DOLGU, 11: AI_DOLGU,
                           12: T.KIRMIZI if s["soz"] else AI_DOLGU, 15: T.SARI})
        if ws.max_row == 1:
            T.satir(ws, ["—", None, "", "Onay gerektiren e-posta yok"])
    o = T.sayfa(wb, "Özet", ["Gösterge", "Adet"], (34, 10))
    for p in ("P1", "P2", "P3"):
        T.satir(o, [f"Öncelik {p}", sum(s["oncelik"] == p for s in sonuc.values())])
    T.satir(o, ["İnsan onayı gerekli", sum(s["onay"] for s in sonuc.values())])
    T.satir(o, ["Taslağında söz verme bulunan", sum(bool(s["soz"]) for s in sonuc.values())])
    T.satir(o, [])
    for k, n in Counter(s["kategori"] for s in sonuc.values()).most_common():
        T.satir(o, [k, n])
    T.satir(o, [])
    T.satir(o, ["Model", llm.kullanim_ozeti()])
    T.satir(o, ["Renkler", "Mor: yapay zekâ; kırmızı taslak: söz verme bulundu, göndermeden düzeltin; sarı: onay gerekli"])
    T.kaydet(wb, cikti)


def ozet(epostalar, sonuc) -> dict:
    gunluk = {}
    for e in epostalar:
        s = sonuc.get(e["id"])
        if not s or not e["zaman"]:
            continue
        g = gunluk.setdefault(e["zaman"].date().isoformat(), {"talep": 0, "p1": 0, "onay": 0})
        g["talep"] += 1
        g["p1"] += s["oncelik"] == "P1"
        g["onay"] += s["onay"]
    return {"gunluk": dict(sorted(gunluk.items())), "kategoriler": dict(Counter(s["kategori"] for s in sonuc.values()).most_common())}


def calistir(girdi, cikti: Path, sla_yolu: Path | None = None, urun: str = "abonelikli mobil uygulama", evet: bool = False) -> dict:
    epostalar = epostalari_oku(girdi)
    if not epostalar:
        raise llm.LLMHatasi("Okunacak e-posta yok.")
    sla = dict(VARSAYILAN_SLA)
    if sla_yolu:
        sla.update({k: v for k, v in json.loads(sla_yolu.read_text(encoding="utf-8")).items() if k in VARSAYILAN_SLA})
    for e in epostalar:
        e["kural"] = kurallar(e)
    hassas = sum(bool(set(e["kural"]) & {"iade", "silme", "yasal", "guvenlik"}) for e in epostalar)
    print(f"[OK] {len(epostalar)} e-posta · kodun bulduğu hassas konu: {hassas}")
    llm.onay_al(f"{len(epostalar)} e-postanın konu ve gövdesi (gönderen adresi hariç; e-posta / telefon / IBAN / TCKN maskeli) gönderilecek.", evet)
    sonuc = siniflandir(epostalar, urun)
    duzeltmeler = {e["id"]: kural_uygula(e, sonuc[e["id"]], sla) for e in epostalar if e["id"] in sonuc}
    rapor_yaz(cikti, epostalar, sonuc, duzeltmeler)
    T.ozet_yaz(cikti.parent, PAKET_ADI, ozet(epostalar, sonuc))
    return {"epostalar": epostalar, "sonuc": sonuc, "duzeltmeler": duzeltmeler}


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    try:
        from dotenv import load_dotenv
        load_dotenv(BURASI / ".env")
        load_dotenv()
    except ImportError:
        pass
    p = argparse.ArgumentParser(description="Uygulama destek e-postalarını sınıflandırır, önceliklendirir, hassas talepleri onaya ayırır ve cevap taslağı yazar.")
    p.add_argument("--girdi", nargs="+", type=Path, help="E-posta dökümü: .csv / .xlsx (Tarih, Gönderen, Konu, Mesaj), .mbox ya da .eml klasörü")
    p.add_argument("--sla", type=Path, help="İlk yanıt hedefleri, saat (JSON: P1, P2, P3)")
    p.add_argument("--urun", default="abonelikli mobil uygulama", help="Modele verilecek kısa ürün tanımı (ör. 'fal ve burç uygulaması, Plus aboneliği')")
    p.add_argument("--cikti", type=Path, default=Path("cikti") / "destek_epostalari.xlsx")
    p.add_argument("--evet", action="store_true", help="Veri gönderim onayını sormadan devam et")
    a = p.parse_args(argv)
    if not a.girdi:
        a.girdi = [ORNEK / "destek_epostalari.csv"]
        a.sla = a.sla or ORNEK / "sla.json"
    try:
        s = calistir(a.girdi, a.cikti, a.sla, a.urun, a.evet)
    except (llm.LLMHatasi, ValueError) as h:
        print(f"[X] {h}")
        return 1
    say = Counter(x["oncelik"] for x in s["sonuc"].values())
    onay = sum(x["onay"] for x in s["sonuc"].values())
    print(f"[OK] {len(s['sonuc'])}/{len(s['epostalar'])} e-posta · P1 {say['P1']} · P2 {say['P2']} · P3 {say['P3']} · insan onayı gerekli {onay}")
    for i, d in s["duzeltmeler"].items():
        for x in d:
            if "söz" in x:
                print(f"[!] {i}: {x}")
    print(f"[OK] Rapor: {a.cikti.resolve()}")
    print(f"[i] Kullanım: {llm.kullanim_ozeti()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

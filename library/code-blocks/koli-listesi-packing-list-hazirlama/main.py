"""
Koli Listesi (Packing List) Hazırlama — Workers / Workless kod bloğu
Tekstil ve Konfeksiyon › Ütü-Paket ve Sevkiyat › Sevkiyat Sorumlusu

Sevk edilecek adetlerden (sipariş × model × renk × beden) koli planı ve ihracata uygun packing list çıkarır:
  - Paketleme yöntemi: "solid" (her kolide tek renk-tek beden) veya "asorti" (her kolide sabit beden oranı,
    ör. S1-M2-L2-XL1); asortiye girmeyen kalanlar solid, onlar da koli içini doldurmuyorsa "karışık koli" olur.
  - Koli numaraları sürekli verilir (1–20, 21–35 …); her satır için koli adedi, koli içi adet, toplam adet.
  - Net ağırlık (adet ağırlığı × adet), brüt ağırlık (+ koli darası), koli ölçüsünden hacim (m³) ve toplamlar.
  - Etiket listesi: her koli için bir satır (koli no, içerik, ağırlık) — koli etiketi basımına hazır.
İnternete bağlanmaz.

Kullanım:
    python main.py                                         # örnek sevkiyatla dener
    python main.py --sevk sevk.xlsx --yontem asorti --asorti S=1,M=2,L=2,XL=1 --koli-ici 24
    python main.py --sevk sevk.xlsx --yontem solid --koli-ici 30 --koli-olcu 60x40x40 --dara 0,8
"""
from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from collections import OrderedDict, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BURASI = Path(__file__).resolve().parent


def kucuk(s) -> str:
    return str(s or "").replace("İ", "i").replace("I", "ı").lower().strip()


def katla(s) -> str:
    s = kucuk(s)
    for a, b in zip("ıöüşğç", "iousgc"):
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def sayi(x, varsayilan=0.0):
    if x in (None, ""):
        return varsayilan
    if isinstance(x, (int, float)):
        return float(x)
    try:
        return float(str(x).strip().replace(".", "").replace(",", ".") if "," in str(x) else str(x).strip())
    except ValueError:
        return varsayilan


def tablo_oku(yol: Path) -> list[list]:
    if yol.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(yol, data_only=True, read_only=True)
        try:
            satirlar = [list(r) for r in wb.active.iter_rows(values_only=True)]
        finally:
            wb.close()
    else:
        metin = None
        for kod in ("utf-8-sig", "cp1254"):
            try:
                metin = yol.read_text(encoding=kod)
                break
            except UnicodeDecodeError:
                continue
        ilk = "\n".join(metin.splitlines()[:10])
        satirlar = list(csv.reader(metin.splitlines(), delimiter=max(";\t,", key=ilk.count)))
    return [r for r in satirlar if any(c not in (None, "") for c in r)]


def sevk_oku(yol: Path) -> tuple["OrderedDict[tuple, OrderedDict[str, int]]", dict]:
    """(sipariş, model, renk) → {beden: adet}; adet ağırlığı (kg) model bazında."""
    s = tablo_oku(yol)
    kb = [katla(x) for x in s[0]]
    bul = lambda *a: next((kb.index(katla(x)) for x in a if katla(x) in kb), None)  # noqa: E731
    i_s, i_m, i_r = bul("sipariş no", "po", "po no", "sipariş"), bul("model", "stil", "artikel", "style"), bul("renk", "color", "colour")
    i_b, i_a = bul("beden", "size"), bul("adet", "miktar", "qty", "quantity")
    i_w = bul("adet ağırlığı kg", "birim ağırlık kg", "adet ağırlığı", "birim ağırlık")
    if i_m is None:
        raise SystemExit(f"Sevk dosyasında 'Model' sütunu gerekli. Başlıklar: {s[0]}")
    gr: OrderedDict[tuple, OrderedDict[str, int]] = OrderedDict()
    agirlik = {}
    al = lambda r, i: r[i] if i is not None and i < len(r) else None  # noqa: E731
    if i_b is not None and i_a is not None:
        for r in s[1:]:
            if not al(r, i_m):
                continue
            k = (str(al(r, i_s) or "-"), str(al(r, i_m)), str(al(r, i_r) or "-"))
            gr.setdefault(k, OrderedDict())[str(al(r, i_b)).strip()] = gr.get(k, {}).get(str(al(r, i_b)).strip(), 0) + int(sayi(al(r, i_a)))
            if i_w is not None and al(r, i_w):
                agirlik[k[1]] = sayi(al(r, i_w))
        return gr, agirlik
    sabit = {i_s, i_m, i_r, i_w, bul("toplam", "total")}
    bedenler = [(i, str(b).strip()) for i, b in enumerate(s[0]) if i not in sabit and str(b or "").strip()]
    for r in s[1:]:
        if not al(r, i_m):
            continue
        k = (str(al(r, i_s) or "-"), str(al(r, i_m)), str(al(r, i_r) or "-"))
        gr[k] = OrderedDict((b, int(sayi(al(r, i)))) for i, b in bedenler if sayi(al(r, i)))
        if i_w is not None and al(r, i_w):
            agirlik[k[1]] = sayi(al(r, i_w))
    return gr, agirlik


# ----------------------------------------------------------------------------
# Koli planı
# ----------------------------------------------------------------------------

def koli_plani(gruplar, yontem: str, koli_ici: int, asorti: dict[str, int] | None) -> list[dict]:
    """Dönüş: koli tipleri listesi — her biri {anahtar, icerik {beden: adet/koli}, koli, tur}."""
    satirlar = []
    for k, bedenler in gruplar.items():
        kalan = dict(bedenler)
        if yontem == "asorti" and asorti:
            paket = sum(asorti.values())
            kat = max(1, koli_ici // paket)                     # kolide kaç asorti seti var
            icerik = {b: n * kat for b, n in asorti.items()}
            if all(b in kalan for b in asorti):
                koli = min(kalan[b] // icerik[b] for b in asorti if icerik[b])
                if koli:
                    satirlar.append({"anahtar": k, "icerik": icerik, "koli": koli, "tur": "Asorti"})
                    for b in asorti:
                        kalan[b] -= icerik[b] * koli
        karisik = {}
        for b, a in kalan.items():
            if a <= 0:
                continue
            tam = a // koli_ici
            if tam:
                satirlar.append({"anahtar": k, "icerik": {b: koli_ici}, "koli": tam, "tur": "Solid"})
            if a % koli_ici:
                karisik[b] = a % koli_ici
        # Kalanlar: koli içini dolduracak şekilde karışık kolilere (beden sırasıyla)
        tampon, koliler = {}, []
        for b, a in karisik.items():
            while a:
                bos = koli_ici - sum(tampon.values())
                al = min(bos, a)
                tampon[b] = tampon.get(b, 0) + al
                a -= al
                if sum(tampon.values()) == koli_ici:
                    koliler.append(tampon)
                    tampon = {}
        if tampon:
            koliler.append(tampon)
        for ic in koliler:
            tur = "Solid (eksik)" if len(ic) == 1 else "Karışık"
            if satirlar and satirlar[-1]["anahtar"] == k and satirlar[-1]["icerik"] == ic and satirlar[-1]["tur"] == tur:
                satirlar[-1]["koli"] += 1
            else:
                satirlar.append({"anahtar": k, "icerik": ic, "koli": 1, "tur": tur})
    return satirlar


def numarala(satirlar: list[dict], agirlik: dict, varsayilan_agirlik: float, dara: float, olcu: tuple[float, float, float] | None):
    no = 1
    hacim = (olcu[0] * olcu[1] * olcu[2] / 1_000_000) if olcu else None
    for s in satirlar:
        s["bas"], s["son"] = no, no + s["koli"] - 1
        no += s["koli"]
        s["koli_adet"] = sum(s["icerik"].values())
        s["toplam"] = s["koli_adet"] * s["koli"]
        w = agirlik.get(s["anahtar"][1], varsayilan_agirlik)
        s["net_koli"] = round(w * s["koli_adet"], 3)
        s["brut_koli"] = round(s["net_koli"] + dara, 3)
        s["net"] = round(s["net_koli"] * s["koli"], 2)
        s["brut"] = round(s["brut_koli"] * s["koli"], 2)
        s["hacim"] = round(hacim * s["koli"], 3) if hacim else None


def calistir(sevk_yolu: Path, cikti: Path, yontem: str = "solid", koli_ici: int = 24, asorti: dict[str, int] | None = None,
             adet_agirlik: float = 0.25, dara: float = 0.8, olcu: tuple[float, float, float] | None = (60, 40, 40),
             gonderici: str = "", alici: str = "") -> dict:
    gruplar, agirlik = sevk_oku(sevk_yolu)
    if yontem == "asorti" and not asorti:
        raise SystemExit("--yontem asorti için --asorti verilmeli (ör. S=1,M=2,L=2,XL=1).")
    if asorti and koli_ici % sum(asorti.values()):
        raise SystemExit(f"Koli içi adet ({koli_ici}) asorti seti büyüklüğünün ({sum(asorti.values())}) katı olmalı.")
    satirlar = koli_plani(gruplar, yontem, koli_ici, asorti)
    numarala(satirlar, agirlik, adet_agirlik, dara, olcu)
    toplam_sevk = sum(sum(v.values()) for v in gruplar.values())
    paketlenen = sum(s["toplam"] for s in satirlar)
    if toplam_sevk != paketlenen:
        raise SystemExit(f"İç hata: sevk {toplam_sevk} ≠ paketlenen {paketlenen}")
    _rapor(gruplar, satirlar, yontem, koli_ici, asorti, dara, olcu, gonderici, alici, cikti)
    return {"gruplar": gruplar, "satirlar": satirlar}


# ----------------------------------------------------------------------------
# Rapor
# ----------------------------------------------------------------------------

BASLIK_DOLGU = PatternFill("solid", fgColor="1D1D1F")
BASLIK_YAZI = Font(bold=True, color="FFFFFF")
KARISIK = PatternFill("solid", fgColor="FFF4CE")


def _baslik(ws, satir=1):
    for c in ws[satir]:
        c.fill, c.font = BASLIK_DOLGU, BASLIK_YAZI


def icerik_metni(ic: dict[str, int]) -> str:
    return " ".join(f"{b}:{n}" for b, n in ic.items())


def _rapor(gruplar, satirlar, yontem, koli_ici, asorti, dara, olcu, gonderici, alici, cikti):
    bedenler = list(OrderedDict.fromkeys(b for v in gruplar.values() for b in v))
    wb = Workbook()
    p = wb.active
    p.title = "Packing List"
    p.append(["PACKING LIST / KOLİ LİSTESİ"])
    p["A1"].font = Font(bold=True, size=13)
    if gonderici or alici:
        p.append([f"Gönderici / Shipper: {gonderici}", "", "", "", f"Alıcı / Consignee: {alici}"])
    p.append([])
    bas = ["Koli No / Carton No", "Koli / Ctns", "Sipariş / PO", "Model / Style", "Renk / Color", "Tür"] + bedenler + \
          ["Koli İçi / Pcs per Ctn", "Toplam Adet / Total Pcs", "Net kg / Koli", "Brüt kg / Koli", "Toplam Net kg", "Toplam Brüt kg", "Hacim m³ / CBM"]
    p.append(bas)
    bs = p.max_row
    _baslik(p, bs)
    for s in satirlar:
        no = f"{s['bas']}" if s["bas"] == s["son"] else f"{s['bas']}–{s['son']}"
        sip, model, renk = s["anahtar"]
        p.append([no, s["koli"], sip, model, renk, s["tur"]] + [s["icerik"].get(b) for b in bedenler] +
                 [s["koli_adet"], s["toplam"], s["net_koli"], s["brut_koli"], s["net"], s["brut"], s["hacim"]])
        if s["tur"] != "Asorti" and s["tur"] != "Solid":
            p.cell(p.max_row, 6).fill = KARISIK
    n = len(bedenler)
    toplam = ["TOPLAM / TOTAL", sum(s["koli"] for s in satirlar), "", "", "", ""] + \
             [sum(s["icerik"].get(b, 0) * s["koli"] for s in satirlar) for b in bedenler] + \
             ["", sum(s["toplam"] for s in satirlar), "", "", round(sum(s["net"] for s in satirlar), 2),
              round(sum(s["brut"] for s in satirlar), 2), round(sum(s["hacim"] or 0 for s in satirlar), 3) if olcu else None]
    p.append(toplam)
    for c in p[p.max_row]:
        c.font = Font(bold=True)
    for j, w in enumerate([14, 9, 14, 12, 12, 13] + [6] * n + [11, 12, 11, 11, 12, 12, 11], 1):
        p.column_dimensions[get_column_letter(j)].width = w
    p.freeze_panes = p.cell(bs + 1, 3)

    o = wb.create_sheet("Özet")
    o.append(["Sipariş", "Model", "Renk", "Sevk Adedi", "Koli", "Net kg", "Brüt kg"])
    _baslik(o)
    ozet = defaultdict(lambda: [0, 0, 0.0, 0.0])
    for s in satirlar:
        x = ozet[s["anahtar"]]
        x[0] += s["toplam"]
        x[1] += s["koli"]
        x[2] += s["net"]
        x[3] += s["brut"]
    for (sip, model, renk), (adet, koli, net, brut) in ozet.items():
        o.append([sip, model, renk, adet, koli, round(net, 2), round(brut, 2)])
    o.append([])
    o.append(["Yöntem", yontem + (f" ({icerik_metni(asorti)})" if asorti else ""), "Koli içi", koli_ici, "Dara (kg)", dara])
    if olcu:
        o.append(["Koli ölçüsü (cm)", "×".join(f"{x:g}" for x in olcu)])
    for j, w in enumerate((16, 14, 12, 12, 8, 10, 10), 1):
        o.column_dimensions[get_column_letter(j)].width = w

    e = wb.create_sheet("Koli Etiketleri")
    e.append(["Koli No", "Toplam Koli", "Sipariş / PO", "Model", "Renk", "İçerik", "Adet", "Net kg", "Brüt kg"])
    _baslik(e)
    toplam_koli = sum(s["koli"] for s in satirlar)
    for s in satirlar:
        for no in range(s["bas"], s["son"] + 1):
            sip, model, renk = s["anahtar"]
            e.append([no, toplam_koli, sip, model, renk, icerik_metni(s["icerik"]), s["koli_adet"], s["net_koli"], s["brut_koli"]])
    for j, w in enumerate((8, 11, 14, 12, 12, 30, 7, 8, 8), 1):
        e.column_dimensions[get_column_letter(j)].width = w
    e.freeze_panes = "B2"
    cikti.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cikti)


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ornek = BURASI / "ornek_veri"
    ap = argparse.ArgumentParser(description="Sevk adetlerinden koli planı ve ihracata uygun packing list hazırlar.")
    ap.add_argument("--sevk", type=Path, default=ornek / "sevk.csv",
                    help="Sevk adetleri: Sipariş No, Model, Renk, S, M, L … (veya Beden, Adet) [, Adet Ağırlığı (kg)]")
    ap.add_argument("--yontem", choices=["solid", "asorti"], default="solid", help="Paketleme: solid (tek beden) veya asorti (beden oranı)")
    ap.add_argument("--asorti", help="Asorti seti, ör. S=1,M=2,L=2,XL=1")
    ap.add_argument("--koli-ici", type=int, default=24, help="Koli içi adet (varsayılan 24)")
    ap.add_argument("--agirlik", type=float, default=0.25, help="Sevk dosyasında yoksa adet ağırlığı, kg (varsayılan 0,25)")
    ap.add_argument("--dara", type=float, default=0.8, help="Boş koli ağırlığı, kg (varsayılan 0,8)")
    ap.add_argument("--koli-olcu", default="60x40x40", help="Koli ölçüsü, cm (varsayılan 60x40x40; 'yok' = hacim hesaplanmaz)")
    ap.add_argument("--gonderici", default="", help="Gönderici (shipper) adı")
    ap.add_argument("--alici", default="", help="Alıcı (consignee) adı")
    ap.add_argument("--cikti", type=Path, default=Path("cikti") / "packing_list.xlsx")
    a = ap.parse_args(argv)
    asorti = None
    if a.asorti:
        asorti = OrderedDict()
        for parca in a.asorti.split(","):
            b, _, n = parca.partition("=")
            asorti[b.strip()] = int(sayi(n))
    elif a.sevk == ornek / "sevk.csv" and a.yontem == "solid":
        a.yontem, asorti = "asorti", OrderedDict([("S", 1), ("M", 2), ("L", 2), ("XL", 1)])
    olcu = None
    if katla(a.koli_olcu) not in ("yok", ""):
        m = re.match(r"^\s*([\d.,]+)\s*[xX×*]\s*([\d.,]+)\s*[xX×*]\s*([\d.,]+)\s*$", a.koli_olcu)
        if not m:
            raise SystemExit("--koli-olcu biçimi 60x40x40 olmalı")
        olcu = tuple(sayi(x) for x in m.groups())
    s = calistir(a.sevk, a.cikti, a.yontem, a.koli_ici, asorti, a.agirlik, a.dara, olcu, a.gonderici, a.alici)
    koli = sum(x["koli"] for x in s["satirlar"])
    adet = sum(x["toplam"] for x in s["satirlar"])
    karisik = sum(x["koli"] for x in s["satirlar"] if x["tur"] not in ("Asorti", "Solid"))
    print(f"[OK] {adet} adet · {koli} koli (karışık/eksik {karisik}) · brüt {sum(x['brut'] for x in s['satirlar']):.1f} kg"
          + (f" · {sum(x['hacim'] for x in s['satirlar']):.2f} m³" if olcu else ""))
    print(f"[OK] Rapor: {a.cikti.resolve()}")


if __name__ == "__main__":
    sys.exit(main())

# Alacak Yaşlandırma · Kod Bloğu

> Finans › Finans Uzmanı · Workers / Workless

Muhasebe programından alınan **cari hesap hareketlerinden** (borç/alacak) tahsilatları en eski faturadan
başlayarak kapatır (**FIFO**) ve açık kalan faturaları vadeye göre yaşlandırır. Müşteri bazında ağırlıklı
ortalama gecikme gününü, kredi limiti kullanımını ve riskli alacakları gösterir. İnternete bağlanmaz.

## Dilimler

Vadesi gelmemiş · 1-30 · 31-60 · 61-90 · 91-180 · 180+ gün (rapor tarihi − vade tarihi).
Vade tarihi olmayan faturalarda belge tarihi + `--vade` gün kullanılır.

## İşaretlenen riskler

- 90 günü aşan alacak
- 180 günü aşan alacak (şüpheli alacak değerlendirmesi önerilir; VUK md. 323 koşulları ayrıca aranır)
- Kredi limiti aşımı (limit dosyası verilirse)
- Faturaları aşan tahsilat (avans / alacaklı bakiye) — Bilgi sayfasında uyarı

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                 # örnek veriyle dener (rapor tarihi 30.09.2026)
python main.py --girdi cari_hareketler.xlsx --tarih 30.09.2026
python main.py --girdi cari_hareketler.xlsx --vade 30 --limitler kredi_limitleri.xlsx
```

**Girdi (hareket listesi):** Cari (veya Cari Kodu), Tarih, Borç, Alacak zorunlu; Vade Tarihi, Belge No isteğe bağlı.
**Girdi (açık kalem listesi):** Cari, Tarih, Kalan zorunlu; Vade Tarihi, Tutar, Belge No isteğe bağlı.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

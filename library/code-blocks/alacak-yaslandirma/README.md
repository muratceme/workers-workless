# Alacak Yaşlandırma · Kod Bloğu

> Finans › Finans Uzmanı · Gerisi Bende

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

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 213 sayılı Vergi Usul Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.4.213.pdf) | Resmî metinle karşılaştırılan maddeler: 274, 280, 315, mükerrer 315, 320, 323 |

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.

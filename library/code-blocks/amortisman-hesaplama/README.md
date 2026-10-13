# Amortisman Hesaplama · Kod Bloğu

> Muhasebe › Genel Muhasebe Uzmanı · Workers / Workless

Demirbaş listesinden faydalı ömür ve yönteme göre **seçilen yılın amortismanını, birikmiş amortismanı ve net defter
değerini** hesaplar; her demirbaş için yıl yıl amortisman planı çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Yöntem | Kural |
|---|---|
| Normal | Maliyet ÷ faydalı ömür; her yıl eşit |
| Azalan bakiyeler | Oran = min(2 ÷ faydalı ömür, %50); yıl amortismanı = net defter değeri × oran; son yıl kalan tamamı |
| Aktife giriş yılı | Tam yıl amortisman ayrılır |
| Kıst | İlk yıl = yıllık tutar × (giriş ayı dahil yıl sonuna kalan ay) ÷ 12; ayrılmayan kısım son yıla eklenir |

Yıllık tutarlar kuruşa yuvarlanır; son yıl, toplam tam maliyete eşit olacak şekilde kapatılır.

**Örnek:** 100.000 maliyet, 5 yıl, azalan bakiyeler → oran %40: 40.000 · 24.000 · 14.400 · 8.640 · 12.960.
Kıst örneği: 600.000, 5 yıl, Eylül'de aktife girmiş → ilk yıl 120.000 × 4 ÷ 12 = 40.000; son yıl 120.000 + 80.000.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, 2026 yılı
python main.py --demirbaslar demirbaslar.xlsx
python main.py --demirbaslar demirbaslar.xlsx --yil 2026
```

Zorunlu sütunlar: Demirbaş, Aktife Giriş Tarihi, Maliyet, Faydalı Ömür (yıl). Diğerleri: Kod, Yöntem (boşsa
Normal), Kıst (Evet). `--yil` verilmezse içinde bulunulan yıl hesaplanır.

## Çıktı

`cikti/amortisman.xlsx`:

- `Dönem Amortismanı`: demirbaş bazında yıl amortismanı, birikmiş amortisman, net defter değeri, durum; toplamlar.
- `Amortisman Planı`: her demirbaşın aktife girişten ömür sonuna kadar yıllık planı.
- `Uyarılar`: hesap yılından sonra aktife girenler, kıst ile ilgili notlar, okunamayan satırlar.

## Dikkat

- **Faydalı ömrü ve yöntemi siz verirsiniz.** Kod faydalı ömür listelerini bilmez; varlığın ömrünü ve seçilebilecek
  yöntemi yürürlükteki vergi mevzuatına göre siz belirlersiniz. Yöntem değişikliği kuralları denetlenmez.
- **Kıst amortisman her varlığa uygulanmaz.** Vergi uygulamasında kıst esası belirli varlıklar (örneğin binek
  otomobiller) içindir; "Kıst" sütununu buna göre işaretleyin.
- **Kıst ile azalan bakiyeler birlikte kullanıldığında** kod ertelenen kısmı son yıla ekler ve uyarır; bu birleşimi
  meslek mensubunuzla doğrulayın.
- **Kapsam dışı:** yeniden değerleme, enflasyon düzeltmesi, doğrudan gider yazma sınırı, binek otomobillerde gider
  kısıtlaması, yıl içinde satış veya hurdaya ayırma, ara dönem (aylık, üç aylık) amortisman, TFRS'ye göre kalıntı
  değer ve bileşen yaklaşımı.
- **Muhasebe kaydı üretilmez.** Çıktı hesap tablosudur; kayıt ve beyan sizin sorumluluğunuzdadır.
- **Örnek veri:** Demirbaşlar ve tutarlar kurgusaldır.

## Testler

- Normal ve azalan bakiyeler planı, %50 oran sınırı, son yılın kapatılması, kuruş yuvarlaması.
- Kıst amortisman (yıl başı, yıl sonu, tek yıllık ömür, azalan bakiyelerle birlikte).
- Yıl seçimi (aktife girmemiş, devam ediyor, son yıl, tamamen amorti edilmiş), toplamlar; okuma ve komut satırı hataları.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

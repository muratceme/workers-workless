# Mükellef Beyanname Takvimi · Kod Bloğu

> Beyanname ve Vergi › Serbest Muhasebeci Mali Müşavir (SMMM) · Workers / Workless

Mükelleflerin yükümlülüklerinden **seçilen ayın beyanname ve ödeme takvimini** çıkarır; verilenleri, yaklaşanları ve
gecikenleri mükellef bazında izler. İnternete bağlanmaz.

## Nasıl çalışır?

Beyanname türlerini ve günlerini **siz tanımlarsınız**; kod bu tanımı mükellef listesine uygular.

| Sütun | Anlamı | Örnek |
|---|---|---|
| Periyot | Aylık, 3 Aylık veya Yıllık | Aylık |
| Dönem Sonu Ayları | Boşsa periyottan gelir (3 Aylık: 3, 6, 9, 12; Yıllık: 12) | 3,6,9,12 |
| Kaç Ay Sonra | Dönem bittikten kaç ay sonra verildiği (boşsa 1) | 2 |
| Beyan Günü / Ödeme Günü | Ayın günü veya `son` | 28 |

- Takvim ayından "kaç ay sonra" kadar geri gidilen ay, türün dönem sonu aylarından biriyse beyanname o ay verilir.
- Son gün cumartesi, pazar veya tatil dosyanızdaki bir güne rastlarsa izleyen ilk iş gününe kayar.
- Durum: **Verildi**, **Geç verildi**, **Gecikmiş**, **Yaklaşıyor** (`--yaklasan` gün veya daha az), **Bekliyor**.
  Ödeme günü tanımlıysa ödeme için de aynı durumlar izlenir.

**Örnek:** "Form Ba-Bs; Aylık; kaç ay sonra 1; beyan günü son" → Ekim 2026 takviminde Eylül dönemi; ayın son günü
31.10.2026 cumartesi olduğundan son gün 02.11.2026 pazartesi.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, Ekim 2026, bugün 26.10.2026
python main.py --yukumlulukler yukumlulukler.xlsx --turler beyanname_turleri.xlsx --ay 2026-10
python main.py --yukumlulukler yukumlulukler.xlsx --turler beyanname_turleri.xlsx --durum durum.xlsx --tatiller tatiller.xlsx --ay 2026-10 --yaklasan 5
```

- Yükümlülükler: Mükellef, Beyanname, Sorumlu.
- Durum (isteğe bağlı): Mükellef, Beyanname, Dönem (`2026-09`, `09.2026`, `Eylül 2026`), Verildi, Ödendi (tarih).
- Tatiller (isteğe bağlı): Tarih, Açıklama.

## Çıktı

`cikti/beyanname_takvimi.xlsx`:

- `Takvim`: son gün sırasıyla beyannameler; mükellef sayısı, verilen, kalan, geciken.
- `Mükellef Bazında`: her mükellefin o ayki beyannameleri, durumu ve ödeme durumu.
- `Uyarılar`: gecikmiş ve yaklaşan beyannameler, gecikmiş ödemeler, tanımsız türler, okunamayan satırlar.

## Dikkat

- **Kod vergi takvimini bilmez.** Örnek dosyadaki günler örnektir. Beyan ve ödeme süreleri mevzuatla değişir; GİB sık sık
  süre uzatır. Her ay GİB vergi takviminden doğrulayıp türler dosyanızı güncel tutun.
- **Resmi tatilleri siz girersiniz.** Kod yalnız hafta sonunu bilir; dini bayramlar, yarım günler ve idari izinler
  için tatiller dosyasını doldurun. Yarım gün tatillerde sürenin nasıl işlediğini ayrıca kontrol edin.
- **Özel durumlar kapsam dışıdır.** Özel hesap dönemi, tasfiye, işe başlama ve bırakma, mücbir sebep hâli için
  tarihleri elle yönetin.
- **"Verildi" kaydı sizin beyanınızdır.** Kod tahakkuk fişini veya ödemeyi doğrulamaz.
- **Örnek veri:** Mükellefler ve personel kurgusaldır.

## Testler

- Aylık, üç aylık ve yıllık türlerin doğru aya düşmesi; yıl geçişi; ayın son günü.
- Hafta sonu ve tatil kaydırması; verildi, geç verildi, gecikmiş, yaklaşıyor durumları; ödeme durumu.
- Tanımsız tür, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

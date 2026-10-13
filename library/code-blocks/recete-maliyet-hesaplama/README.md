# Reçete Maliyet Hesaplama · Kod Bloğu

> Ar-Ge ve Ürün Geliştirme › Ar-Ge Uzmanı · Workers / Workless

Reçete bileşen miktarları, hammadde fiyatları ve fire oranlarından **kilogram ve ambalaj başına hammadde maliyetini**
hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Bileşen oranı % | Miktar ÷ reçete toplamı × 100 |
| Brüt ihtiyaç | Miktar ÷ (1 − fire % ÷ 100) |
| Bileşen maliyeti | Brüt ihtiyaç × birim fiyat |
| Mamul miktarı | Reçete toplamı × proses verimi % ÷ 100 |
| Kg maliyeti | Σ bileşen maliyeti ÷ mamul miktarı |
| Ambalaj başına | Kg maliyeti × net ağırlık (g) ÷ 1000 + ambalaj malzemesi maliyeti |
| Maliyet payı % | Bileşen maliyeti ÷ Σ bileşen maliyeti × 100 |

- **Fire %**, bileşenin hazırlanırken kaybıdır (ayıklama, eleme, kap dibinde kalan). Reçeteye net miktar girer; satın
  alınması gereken miktar brüt ihtiyaçtır.
- **Proses verimi %**, karışımın mamule dönüşme oranıdır (pişirme, buharlaşma, hat kaybı). Verilmezse %100 alınır.

**Örnek:** 40 kg salça, fire %2, fiyat 60 TL/kg → brüt 40 ÷ 0,98 = 40,816 kg → 2.448,98 TL. Reçete toplamı 100 kg,
hammadde maliyeti 3.176,27 TL, proses verimi %95 → mamul 95 kg → kg maliyeti 33,43 TL. 500 g ambalaj: 16,72 TL
hammadde + 6,50 TL ambalaj = 23,22 TL.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --recete recete.xlsx --fiyatlar fiyatlar.xlsx
python main.py --recete recete.xlsx --fiyatlar fiyatlar.xlsx --urunler urunler.xlsx
```

- Reçete: Ürün, Bileşen, Miktar, Fire % (her bileşen bir satır; bir dosyada birden çok ürün olabilir).
- Fiyatlar: Bileşen, Birim Fiyat.
- Ürünler (isteğe bağlı): Ürün, Proses Verimi %, Ambalaj Net (g), Ambalaj Maliyeti.

Miktarlarda binlik ayırıcı kullanmayın (`1250,5`).

## Çıktı

`cikti/recete_maliyet.xlsx`:

- `Ürün Maliyeti`: hammadde maliyeti, mamul miktarı, kg maliyeti, ambalaj başına maliyet.
- `Bileşen Detayı`: oran, brüt ihtiyaç, maliyet ve maliyet payı.
- `Uyarılar`: fiyatı olmayan bileşenler, okunamayan satırlar.

## Dikkat

- **Yalnız hammadde ve ambalaj malzemesi.** İşçilik, enerji, amortisman, genel üretim gideri, nakliye ve kâr dahil
  değildir; bu rakam satış fiyatı değildir.
- **Birimler tutarlı olmalı.** Miktar kilogramsa fiyat TL/kg olmalıdır. Litreyle alınan hammaddeyi yoğunluğuyla
  kilograma çevirin; kod birim çevirmez.
- **Reçete yüzdeyle de yazılabilir.** Miktar sütununa oranları girerseniz (toplam 100) sonuç değişmez.
- **Fiyatı olmayan bileşen maliyeti düşük gösterir.** Böyle ürünler "Eksik fiyat" olarak işaretlenir.
- **Fiyatın KDV'li mi KDV'siz mi olduğunu siz belirlersiniz.** Kod KDV hesaplamaz; tüm fiyatları aynı esasta girin.
- **Örnek veri:** Ürünler, reçeteler ve fiyatlar kurgusaldır.

## Testler

- Brüt ihtiyaç, bileşen maliyeti, maliyet payı, kg ve ambalaj başına maliyet.
- Ürün bilgisi olmadan (%100 verim, ambalajsız) hesap; fiyatı olmayan ve fiyatı sıfır bileşen.
- Okunamayan reçete, fiyat ve ürün satırları, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

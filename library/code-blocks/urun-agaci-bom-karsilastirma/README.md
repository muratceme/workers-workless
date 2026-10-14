# Ürün Ağacı (BOM) Karşılaştırma · Kod Bloğu

> Ar-Ge ve Ürün Geliştirme › Ar-Ge Mühendisi · Workers / Workless

İki ürün ağacı sürümünü karşılaştırır; **eklenen, çıkarılan ve miktarı, birimi, fiyatı ya da tanımı değişen
kalemleri** listeler, birim maliyet verilmişse değişikliğin maliyet etkisini hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

Kalemler **üst kod + bileşen kodu** ile eşleştirilir (üst kod sütunu yoksa yalnız bileşen kodu).

| Durum | Kural |
|---|---|
| Eklendi | Yalnız yeni sürümde var |
| Çıkarıldı | Yalnız eski sürümde var |
| Birim değişti | Birim farklı; miktar farkı hesaplanmaz |
| Miktar değişti | Birim aynı, miktar farklı; fark = yeni − eski, fark % = fark ÷ eski × 100 |
| Fiyat değişti | Miktar ve birim aynı, birim maliyet farklı |
| Tanım değişti | Yalnız bileşen adı farklı |
| Aynı | Hiçbir fark yok |

Bir kalem birden çok koşulu sağlıyorsa tablodaki sırayla ilk uyan durum yazılır. Kalem maliyeti = miktar × birim
maliyet; maliyet etkisi = yeni kalem maliyeti − eski kalem maliyeti. Birim maliyet bir sürümde yoksa diğer sürümdeki
fiyat kullanılır. Aynı anahtar bir dosyada birden çok satırdaysa miktarlar toplanır.

**Örnek:** Cıvata miktarı 8 adetten 6 adede indi, birim maliyet 2,50 → fark **−2** (−%25), maliyet etkisi **−5,00**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek veri (bom_v1 → bom_v2)
python main.py --eski bom_v1.xlsx --yeni bom_v2.xlsx
python main.py --eski bom_v1.xlsx --yeni bom_v2.xlsx --cikti rapor/degisiklik.xlsx
```

Zorunlu sütunlar: Bileşen Kodu, Miktar. Ondalık miktarları virgülle yazın (`0,25`); `1.000` biçimi bin olarak okunur.

## Çıktı

`cikti/bom_karsilastirma.xlsx`:

- `Değişiklikler`: yalnız farklı kalemler; eski ve yeni miktar, birim, birim maliyet, kalem maliyeti, maliyet etkisi.
- `Tüm Kalemler`: değişmeyenler dahil.
- `Özet`: durum bazında kalem sayısı, eski ve yeni toplam maliyet, maliyet etkisi.
- `Uyarılar`.

## Dikkat

- **Çok seviyeli ağaç patlatılmaz.** Her satır kendi üst koduyla karşılaştırılır; alt gruptaki bir değişikliğin
  mamul başına toplam miktara etkisi (üst seviye miktarlarıyla çarpılarak) hesaplanmaz. Toplam maliyet de
  satırların düz toplamıdır; alt grupların üst seviyedeki miktarı 1'den farklıysa mamul maliyetini göstermez.
- **Kodu değişen kalem "çıkarıldı + eklendi" görünür.** Kod, bir bileşenin başka bir kodla değiştirildiğini
  (muadil, revizyon) anlayamaz.
- **Birim değişince miktar karşılaştırılmaz.** kg → g gibi dönüşümleri kod yapmaz; bu kalemler ayrıca uyarılır.
  Maliyet etkisi, iki sürümde de miktar × birim maliyet doğru yazıldıysa anlamlıdır.
- **Yalnız bu sütunlara bakar.** Fire oranı, operasyon, tedarikçi, revizyon numarası, geçerlilik tarihi ve alternatif
  malzeme gibi alanlardaki değişiklikler görünmez.
- **Kodlar yazıldığı gibi eşleşir.** Büyük-küçük harf ve Türkçe karakter farkı yok sayılır; baştaki sıfırlar,
  boşluk ve tire farkları ayrı kod sayılır.
- **Örnek veri:** Kodlar, bileşenler ve maliyetler kurgusaldır.

## Testler

- Her durumun ataması ve önceliği; fark, fark yüzdesi ve maliyet etkisi; fiyatı tek sürümde olan kalem.
- Üst koda göre eşleştirme (aynı bileşenin farklı üst kodlar altında ayrı kalem sayılması), mükerrer satırların toplanması.
- Miktar yazımları (0.250, 1,5, 1.000), fiyatsız karşılaştırma, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

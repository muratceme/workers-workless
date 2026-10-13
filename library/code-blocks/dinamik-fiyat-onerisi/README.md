# Dinamik Fiyat Önerisi · Kod Bloğu

> Rezervasyon ve Gelir Yönetimi › Gelir Müdürü (Revenue Manager) · Workers / Workless

Doluluk, pick-up hızı ve rakip fiyatından tarih ve oda tipi bazında **kural tabanlı, gerekçesi yazılı** bir fiyat
önerisi üretir. Kurallar şeffaftır ve hepsi parametreyle değişir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural | Varsayılan |
|---|---|---|
| Günlük hız | Son pick-up ÷ `--pickup-gun` | 7 gün |
| Tahmini doluluk % | (Satılan + günlük hız × min(kalan gün, `--projeksiyon`)) ÷ kapasite × 100 | 14 gün |
| Ayar % | (Tahmini doluluk − `--hedef`) × `--duyarlilik`; ±`--azami` ile sınırlı | 85 · 0,5 · 20 |
| Ham fiyat | Mevcut fiyat × (1 + ayar ÷ 100) | |
| Rakip bandı | Artışta rakip × (1 + bant) üstüne, indirimde rakip × (1 − bant) altına çıkılmaz | %15 |
| Yuvarlama | `--yuvarla` katına | 10 |
| Taban / tavan | Satırda verilmişse son fiyat bu aralığa çekilir | |

Rakip bandı öneriyi **sınırlar, ters yöne çevirmez**: talep güçlüyken fiyatınız rakibin çok üstünde diye indirim
önerilmez; "Koru" denir ve ayrıca uyarı yazılır.

**Örnek:** Bugün 12.10; 19.10 gecesi kapasite 100, satılan 80, son 7 günde 14 oda pick-up, fiyat 3.000, rakip 3.200.
Günlük hız 2; tahmini doluluk (80 + 2 × 7) ÷ 100 = %94; ayar (94 − 85) × 0,5 = +%4,5; ham fiyat 3.135; rakip bandı
2.720 – 3.680 içinde; 10'a yuvarlanınca **3.140**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, bugün 12.10.2026
python main.py --durum durum.xlsx
python main.py --durum durum.xlsx --tarih 12.10.2026 --hedef 80 --duyarlilik 0,8 --azami 15 --rakip-bant 10 --yuvarla 50
```

Zorunlu sütunlar: Tarih, Kapasite, Satılan Oda, Mevcut Fiyat. Diğerleri: Oda Tipi, Son 7 Gün Pick-up, Rakip Fiyat,
Taban Fiyat, Tavan Fiyat. Pick-up değeriniz 7 günlük değilse `--pickup-gun` ile belirtin.

## Çıktı

`cikti/fiyat_onerisi.xlsx`:

- `Fiyat Önerisi`: her tarih ve oda tipi için tahmini doluluk, ayar, önerilen fiyat, değişim ve gerekçe; karar için
  boş sütun.
- `Kurallar`: bu çalıştırmada kullanılan parametreler.
- `Uyarılar`: rakip bandının üstündeki fiyatlar, kapasite üstü satış, geçmiş tarihler, okunamayan satırlar.

## Dikkat

- **Gelir yönetimi sistemi değildir.** Talep tahmini yapmaz, fiyat esnekliğini ölçmez, geçmiş yılları bilmez; birkaç
  basit kuralı tutarlı biçimde uygular. Çıktı bir başlangıç önerisidir, karar sizindir.
- **Doğrusal projeksiyon kabadır.** Son günlerin hızının süreceği varsayılır. Rezervasyon eğrisi tarihe yaklaştıkça
  hızlanan veya erken dolan tesislerde tahmin sapar; `--projeksiyon` ve `--hedef` değerlerini kendi eğrinize göre ayarlayın.
- **Varsayılan parametreler bir öneridir.** Hedef doluluk, duyarlılık ve bantlar tesisinize, sezona ve segmente göre
  değişir; kanıtlanmış "doğru" değerler değildir.
- **Rakip fiyatı karşılaştırılabilir olmalı.** Aynı oda tipi, aynı pansiyon, aynı iptal koşulu ve aynı kanal fiyatı
  değilse bant yanıltır.
- **Kontratlı ve sabit fiyatlar kapsam dışıdır.** Acente kontratı, kurumsal anlaşma ve fiyat paritesi yükümlülüklerini
  siz gözetirsiniz.
- **Örnek veri:** Tarihler, doluluklar ve fiyatlar kurgusaldır.

## Testler

- Artış, indirim, azami ayar sınırı, rakip bandının alt ve üst sınırı, taban ve tavan fiyat, yuvarlama yönü.
- Negatif pick-up, dolu kapasite, rakip fiyatı olmayan satır, bandın üstündeki mevcut fiyat; her parametrenin etkisi.
- Geçmiş tarih, sıralama, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

# Karbon Ayak İzi Hesaplama · Kod Bloğu

> Sürdürülebilirlik ve Sosyal Uygunluk › Sosyal Uygunluk Uzmanı · Workers / Workless

Enerji, yakıt ve malzeme tüketim verisinden, **sizin verdiğiniz emisyon faktörleriyle** kapsam 1 ve kapsam 2
emisyonlarını ve ürün başına karbon değerini hesaplar. İnternete bağlanmaz; emisyon faktörü indirmez.

## Nasıl hesaplar?

```
Emisyon (kg CO2e) = tüketim miktarı × emisyon faktörü (kg CO2e ÷ birim)
Ürün başına       = dönemin toplam emisyonu ÷ dönemin üretim miktarı
```

| Kapsam | Ne girer |
|---|---|
| 1 | Tesiste yakılan yakıtlar (doğal gaz, motorin, LPG, kömür ...) ve şirket araçları |
| 2 | Satın alınan elektrik, buhar, ısı |
| 3 | Diğer dolaylı kalemler (ör. hammadde); yalnız girdiğiniz kalemler, yani kısmi |

- Kapsam, faktör dosyasında her kaynak için sizin yazdığınız değerdir (1, 2 veya 3).
- **Birim çevrilmez.** Tüketim birimi faktör birimiyle aynı olmalıdır (kWh ile MWh uyuşmaz; satır alınmaz ve uyarı verilir).
- Faktörü olmayan kaynak hesaba alınmaz ve uyarı olarak listelenir.
- Genel "ürün başına" değeri yalnız üretim miktarı girilen dönemlerden hesaplanır.
- **Örnek:** 10.000 m³ doğal gaz × 2,00 = 20.000 kg; 100.000 kWh elektrik × 0,44 = 44.000 kg; 1.000 L motorin × 2,68
  = 2.680 kg → toplam 66,68 t CO2e; 50.000 adet üretimle 1,3336 kg CO2e / adet. (Faktörler örnek değerlerdir.)

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --tuketimler tuketimler.xlsx --faktorler emisyon_faktorleri.xlsx --uretim uretim.xlsx
```

| Dosya | Sütunlar |
|---|---|
| Tüketimler | Kaynak, Miktar (zorunlu); Dönem (AA.YYYY veya YYYY-AA), Tesis, Birim |
| Emisyon faktörleri | Kaynak, Emisyon Faktörü (kg CO2e / birim), Kapsam (zorunlu); Birim, Kaynakça |
| Üretim (isteğe bağlı) | Dönem, Üretim Miktarı; Tesis, Birim |

## Çıktı

`cikti/karbon_ayak_izi.xlsx`:

- `Özet`: kapsam 1, 2, 3 ve toplam (t CO2e); ürün başına emisyon.
- `Dönem Bazında`: dönem dönem kapsamlar, üretim ve ürün başına değer.
- `Kaynak Bazında`: kaynak başına tüketim, faktör, emisyon ve pay.
- `Tesis Bazında`: tesis başına kapsamlar.
- `Faktörler`: kullanılan faktörler ve kaynakçaları (kaynakçası boş olan işaretlenir).
- `Uyarılar`: faktörü olmayan kaynaklar, birim uyuşmazlıkları, okunamayan satırlar.

## Dikkat

- **Örnek dosyadaki emisyon faktörleri doğrulanmamış örnek değerlerdir.** Raporlamada kullanmayın. Yakıt faktörlerini
  ulusal envanter veya kabul görmüş yayınlardan, elektrik şebeke faktörünü ilgili yılın resmî verisinden ya da
  tedarikçinizden alın; kaynağını ve yılını Kaynakça sütununa yazın.
- **Doğrulanmış bir envanter değildir.** Organizasyon sınırı, soğutucu gaz kaçakları, proses emisyonları, atık,
  nakliye ve satın alınan mal ve hizmetlerin tamamı bu hesabın dışındadır. Müşteri veya mevzuat raporlaması için
  ilgili standardın (ör. GHG Protocol, ISO 14064) gereklerini ve gerekiyorsa üçüncü taraf doğrulamasını uygulayın.
- **Kapsam 3 kısmidir.** Yalnız girdiğiniz kalemleri içerir; "kapsam 3 toplamı" olarak sunulmamalıdır.
- **Ürün başına değer basit bir bölümdür.** Farklı ürünler arasında paylaştırma (ağırlık, süre, ciro) yapılmaz;
  ürün karbon ayak izi (yaşam döngüsü) hesabı değildir.
- **Yenilenebilir elektrik** sertifikaları veya öz tüketim için ayrı kaynak satırı ve kendi faktörünüzü kullanın.
- **Örnek veri:** Tesis, tüketimler ve üretim miktarları kurgusaldır.

## Testler

- Kapsam toplamları, kaynak payları, ürün başına değer.
- Dönem ve tesis kırılımı; üretimi olmayan dönemler.
- Faktörü olmayan kaynak, birim uyuşmazlığı, birim yazım farkları; okuma hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

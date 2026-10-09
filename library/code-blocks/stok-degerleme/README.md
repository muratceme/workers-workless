# Stok Değerleme · Kod Bloğu

> Muhasebe › Maliyet Muhasebesi Uzmanı · Workers / Workless

Stok hareketlerinden dönem sonu stok değerini ve satılan malın (kullanılan malzemenin) maliyetini hesaplar. Üç
yöntemi yan yana gösterir: FIFO, hareketli ağırlıklı ortalama ve dönem sonu ağırlıklı ortalama. Raporun ana yöntemi
`--yontem` ile seçilir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Yöntem | Çıkış maliyeti |
|---|---|
| FIFO (ilk giren ilk çıkar) | En eski giriş katmanından |
| Hareketli ağırlıklı ortalama | Çıkış anındaki ortalama; her girişte ortalama yeniden hesaplanır |
| Dönem sonu ağırlıklı ortalama | (devir + girişler) ÷ miktar; tüm çıkışlar ve kalan stok bu birim maliyetle |

**Hareket türleri:**
- **Devir, Giriş:** Fiyatlı girişlerdir.
- **Çıkış:** Satış, sarf, tüketim veya fire.
- **Alış iadesi:** Kendi fiyatıyla düşülür; FIFO'da önce aynı fiyatlı katmandan alınır.
- **Satış iadesi:** Fiyatsız giriş. FIFO ve hareketli ortalamada son çıkış maliyetiyle geri alınır; dönem sonu
  ortalamada çıkışlardan düşülür.

**Eksi stok:**
- Çıkış mevcut stoktan fazlaysa eksik kısım son maliyetle maliyetlenir.
- Sonraki giriş önce eksiği kapatır. Aradaki fiyat farkı SMM'ye "eksi stok düzeltmesi" olarak yansır.
- Böylece her yöntemde giren değer = SMM + kalan stok olur.

| Kontrol | Önem |
|---|---|
| Eksi stok; dönem sonunda eksi kalan stok | Yüksek |
| Son alış fiyatının dönem ortalamasından %25'ten fazla sapması | Orta |
| Dönemde hiç hareket görmemiş stok | Orta |
| Net gerçekleşebilir değer (tahmini satış fiyatı − satış gideri) birim maliyetin altında | Orta |
| Fiyatsız giriş satış iadesi sayıldı | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 4 stok, Ocak–Mart 2026
python main.py --hareketler stok_hareketleri.xlsx --yontem hareketli
python main.py --hareketler h.xlsx --yontem fifo --tarih 31.12.2026 --satis-fiyatlari fiyatlar.xlsx
```

| Dosya | Sütunlar |
|---|---|
| Hareketler | Tarih, Stok Kodu, Stok Adı, Birim, Hareket, Miktar, Birim Fiyat (veya Tutar), Belge No |
| Satış fiyatları (isteğe bağlı) | Stok Kodu, Tahmini Satış Fiyatı, Satış Gideri |

Aynı gündeki hareketlerde önce girişler işlenir.

## Çıktı

`stok_degerleme.xlsx`:
- `Değerleme Özeti`: stok bazında miktar, birim maliyet, ana yöntemle değer ve SMM; üç yöntemin değeri.
- `Yöntem Karşılaştırması`: toplam stok değeri, SMM, eksi stok düzeltmesi ve fark.
- `Hareket Detayı`: ana yöntemle her hareketin maliyeti ve kalan miktar / değer (FIFO ve hareketli ortalamada).
- `FIFO Katmanları`: dönem sonunda kalan giriş katmanları.
- `Değer Düşüklüğü`: net gerçekleşebilir değer testi; boş "Karar" sütunu.
- `Uyarılar`.

## Dikkat

- **Yöntem seçimi:** Vergi Usul Kanunu'na göre emtia maliyet bedeliyle değerlenir. Kullanılacak yöntem, sürekliliği
  ve değer düşüklüğünün vergisel sonuçları (VUK md. 274 ve 278) için mali müşavirinize danışın. TFRS / BOBİ FRS
  raporlamasında net gerçekleşebilir değer kuralı ayrıca uygulanır.
- **Maliyet kapsamı:** Girişlerdeki birim fiyat, maliyete eklenecek giderleri (nakliye, gümrük vb.) içermelidir.
  Paket bu dağıtımı yapmaz.
- **Örnek veri:** Örnek hareketler ve fiyatlar kurgusaldır.

## Testler

Şunlar test edilir:
- Üç yöntem elle hesaplanmış değerlerle.
- Alış ve satış iadesi; eksi stok düzeltmesi; her yöntemde değer korunumu.
- Değerleme tarihi, net gerçekleşebilir değer testi, uyarılar, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

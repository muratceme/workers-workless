# Fiyat Değişikliği Etiket Listesi · Kod Bloğu

> Mağaza Operasyonları › Mağaza Müdür Yardımcısı · Workers / Workless

Merkezden gelen fiyat değişikliklerini **reyon → raf → sıra** düzeninde, mağazada yürüyerek değiştirilebilecek bir **etiket değişim
listesine** çevirir. İndirimli ürünlerde etikette gösterilecek **önceki fiyatı** (son 30 günün en düşük fiyatı) hesaplar ve indirim
olarak gösterilemeyecek ürünleri işaretler. İnternete ve mağaza yazılımına bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Değişim % | (yeni − eski) ÷ eski × 100; yön: zam / indirim / değişmedi |
| Etikette önceki fiyat (indirim) | Geçerlilik gününden önceki 30 gün (`--gun`) içinde uygulanan en düşük fiyat: fiyat geçmişi dosyası + eski fiyat |
| Çabuk bozulan ürün | Meyve, sebze gibi ürünlerde ("Çabuk Bozulan" = Evet) indirimden bir önceki fiyat |
| İndirim gösterilemez | Yeni fiyat, son 30 günün en düşük fiyatından düşük değilse |
| Sıralama | Reyon, raf ve sıra (sayılar doğal sırayla: Raf 2, Raf 10); bir ürün birden fazla yerde ise her konum için ayrı etiket; konumu olmayanlar sonda |

Dayanak: Fiyat Etiketi Yönetmeliği'nde 18 Şubat 2022'de yapılan ve 1 Mart 2022'de yürürlüğe giren değişiklik. İndirimli satışta
indirimden önceki fiyat olarak son 30 gün içinde uygulanan en düşük fiyat esas alınır; çabuk bozulan ürünlerde indirimden bir önceki
fiyat esas alınır.

**Uyarılar:** indirim olarak gösterilemez · okunamayan satır (Yüksek) · büyük fiyat değişimi (`--esik` %50) · raf konumu yok · aynı ürün iki
kez (Orta) · önceki fiyat eski fiyattan farklı · fiyat değişmemiş · geçerlilik tarihi geçmiş (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (15.10.2026 fiyat değişiklikleri)
python main.py --degisiklikler degisiklikler.xlsx --yerlesim yerlesim.xlsx --gecmis fiyat_gecmisi.xlsx
python main.py --degisiklikler degisiklikler.xlsx --esik 30
```

Fiyat geçmişi verilmezse önceki fiyat olarak yalnız eski fiyat kullanılır; bu durumda son 30 gündeki daha düşük fiyatlar görülemez.

## Çıktı

`cikti/etiket_listesi.xlsx`: `Etiket Listesi` (değiştirildi işareti için boş sütunlu; indirim gösterilemeyenler kırmızı), `Reyon Özeti`, `Uyarılar`.

## Dikkat

- **Mevzuatı doğrulayın.** Fiyat etiketi ve indirimli satış kuralları (ilan, reklam, kampanya süresi bildirimi) değişebilir; güncel
  yönetmelik metnini ve Ticaret Bakanlığı duyurularını kontrol edin.
- **Fiyat geçmişi eksiksiz olmalı.** Son 30 günde uygulanan kampanya fiyatları geçmiş dosyasında yoksa önceki fiyat yanlış hesaplanır.
- **Raf etiketi ile kasa fiyatı aynı olmalı.** Etiketler değiştirildikten sonra kasa sistemindeki fiyatla örnek karşılaştırma yapın.
- **Örnek veri:** Ürünler, fiyatlar ve raf yerleri kurgusaldır.

## Testler

- 30 gün penceresi, çabuk bozulan ürün kuralı, indirim gösterilemez durumu, zam / değişmedi / fiyatı yok, pencere parametresi.
- Doğal sıralama, birden fazla konum, reyon özeti, uyarılar, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

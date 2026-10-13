# Teklif Takip Listesi · Kod Bloğu

> Satış › Satış Destek Uzmanı · Workers / Workless

Verdiğiniz tekliflerin listesinden **bugün aranması gerekenleri** öne alır; geçerliliği dolan ve uzun süredir
takip edilmeyen teklifleri işaretler, dönüşüm oranını temsilci bazında hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Bekleme süresi | Durum tarihi − teklif tarihi (açık teklifler) |
| Takipsiz geçen gün | Durum tarihi − son takip tarihi; son takip yoksa teklif tarihi |
| Açık teklifin durumu | Sırayla: geçerliliği doldu → geçerliliğe `--gecerlilik` (5) gün veya az kaldı → `--takip` (7) gün veya daha uzun süredir aranmadı → bekliyor |
| Dönüşüm oranı % | Kazanılan ÷ (kazanılan + kaybedilen) × 100; adet ve tutar bazında ayrı ayrı |
| Sonuçlanma süresi | Sonuç tarihi − teklif tarihi (kazanılan ve kaybedilen teklifler) |
| Yaş dilimleri | Açık tekliflerin bekleme süresine göre 0–7, 8–15, 16–30, 31+ gün (`--dilimler`) |

İptal edilen ve hâlâ açık olan teklifler dönüşüm oranına girmez.

**Örnek:** 3 teklif kazanıldı, 3 teklif kaybedildi → adet bazında dönüşüm **%50**. Kazanılan 1.915.000, kaybedilen
805.000 → tutar bazında **%70,40**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --teklifler teklifler.xlsx
python main.py --teklifler teklifler.xlsx --takip 5 --gecerlilik 3 --dilimler "10;20;45"
```

Zorunlu sütunlar: Teklif No, Teklif Tarihi, Tutar. Durum boşsa teklif açık sayılır; kabul edilen yazımlar: Açık /
Bekliyor / Gönderildi, Kazanıldı / Onaylandı / Siparişe Dönüştü, Kaybedildi / Reddedildi, İptal.

## Çıktı

`cikti/teklif_takip_listesi.xlsx`:

- `Açık Teklifler`: önce geçerliliği dolanlar, sonra süresi yaklaşanlar ve takip bekleyenler; içinde tutara göre.
- `Özet`: adetler, tutarlar, dönüşüm oranı, ortalama sonuçlanma süresi ve açık tekliflerin yaş dağılımı.
- `Temsilciler`: aynı göstergeler temsilci bazında.
- `Kayıp Nedenleri`: kaybedilen tekliflerin neden dağılımı.
- `Tüm Teklifler` ve `Uyarılar`.

## Dikkat

- **Eşikler yönetim tercihidir.** 7 günde takip ve 5 gün kala geçerlilik uyarısı varsayılandır; satış döngünüzün
  uzunluğuna göre değiştirin.
- **Dönüşüm oranı az sayıda teklifte yanıltır.** Bir temsilcinin üç teklifinden biri sonuçlandıysa oran %0 ya da
  %100 çıkar; adet sütunuyla birlikte okuyun.
- **Tutarlar tek para biriminde olmalı** ve KDV dahil ya da hariç olması bütün satırlarda aynı olmalıdır. Kod kur
  çevirisi yapmaz.
- **Revizyonlar:** Aynı teklifin yeni revizyonunu ayrı satır olarak giriyorsanız eskisini İptal yapın; yoksa hem
  açık tutar hem dönüşüm oranı şişer.
- **Kapsam dışı:** kazanma olasılığıyla ağırlıklandırılmış satış tahmini, teklif kalemleri ve kâr marjı, müşteri
  bazlı geçmiş dönüşüm.
- **Örnek veri:** Müşteriler, kişiler ve tutarlar kurgusaldır.

## Testler

- Açık teklif durumu sınırları: geçerlilik dün / bugün / 5 ve 6 gün sonra; takipsiz 6 ve 7 gün; öncelik sırası.
- Dönüşüm oranı (adet ve tutar), iptal ve açık tekliflerin dışarıda kalması, sonuçlanma süresi, yaş dilimleri.
- Temsilci özeti, kayıp nedeni dağılımı, mükerrer numara, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

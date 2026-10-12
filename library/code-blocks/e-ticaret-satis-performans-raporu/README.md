# E-ticaret Satış Performans Raporu · Kod Bloğu

> Kategori ve İçerik Yönetimi › Kategori Yöneticisi · Workers / Workless

Kanal, kategori ve ürün bazında **ciro, adet, dönüşüm ve iade oranını haftalık** raporlar; önceki haftayla
karşılaştırır. İnternete bağlanmaz; panellerden indirdiğiniz veriyle çalışır.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Hafta | Pazartesi-pazar. Rapor haftası `--hafta` ile verilen günün haftası; verilmezse verideki son hafta |
| Dönüşüm % | Sipariş sayısı ÷ ziyaret × 100 |
| Ortalama sepet | Ciro ÷ sipariş sayısı |
| İade oranı % | İade adedi ÷ satış adedi × 100 |
| Ciro değişimi % | (Bu hafta ciro − önceki hafta ciro) ÷ önceki hafta ciro × 100 |
| Ciro payı % | Ciro ÷ haftanın toplam cirosu × 100 |

| İşaret (ürün bazında) | Koşul |
|---|---|
| İade yüksek | İade oranı > `--iade-esik` (varsayılan %15) |
| Ciro düşüyor | Ciro değişimi ≤ −`--dusus-esik` (varsayılan %30) |
| Bu hafta satış yok | Önceki hafta satılmış, bu hafta hiç kaydı yok |

Aynı haftaya ait birden çok satır (günlük veri) toplanır. Oranlar toplamlardan hesaplanır, satır oranlarının
ortalamasından değil.

**Örnek:** ziyaret 2.000, sipariş 100, adet 120, ciro 36.000, iade 6; önceki hafta ciro 30.000 → dönüşüm %5,
ortalama sepet 360, iade oranı %5, ciro değişimi +%20.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (son hafta: 05–11.10.2026)
python main.py --satislar eticaret_satislari.xlsx --hafta 07.10.2026 --iade-esik 20 --dusus-esik 25
```

Sütunlar: Tarih, Ciro (zorunlu); Kanal, Kategori, Ürün, Ziyaret, Sipariş, Satış Adedi, İade Adedi. Günlük veya
haftalık satırlar verilebilir; karşılaştırma için önceki haftanın satırları da dosyada olmalıdır.

## Çıktı

`cikti/eticaret_satis_raporu.xlsx`:

- `Kanal Bazında`, `Kategori Bazında`: ciro, değişim, pay, adet, sipariş, ziyaret, dönüşüm, ortalama sepet, iade oranı; toplam satırı.
- `Ürün Bazında`: kanal-kategori-ürün kırılımında aynı göstergeler, işaretler ve boş "Aksiyon" sütunu.
- `Uyarılar`: okunamayan satırlar, önceki hafta verisi yoksa bilgi notu.

## Dikkat

- **İade oranı aynı haftanın sayılarındandır.** Bu hafta gelen iadeler çoğunlukla önceki haftaların satışına aittir;
  satışı hızla artan veya azalan üründe oran yanıltır. Kesin oran için iadeyi ait olduğu siparişle eşleştirmek gerekir.
- **Ziyaret tanımı kanala göre değişir.** Web sitesinde oturum, pazaryerinde ürün görüntüleme olabilir; kanallar
  arası dönüşüm karşılaştırması bu yüzden tam değildir.
- **Ciro esası aynı olmalı.** Kanalların hepsinde KDV dahil veya hepsinde hariç; iptaller düşülmüş olmalıdır.
  Komisyon, kargo ve reklam maliyeti bu raporda yoktur; kârlılık göstermez.
- **Bir haftalık değişim oynaktır.** Kampanya, bayram ve stok yokluğu tek haftayı çok etkiler; eğilim için birkaç
  haftayı birlikte değerlendirin.
- **Eşikler varsayılandır.** İade oranı kategoriye göre çok değişir (giyimde yüksek, elektronikte düşük).
- **Örnek veri:** Kanallar, ürünler ve tutarlar kurgusaldır.

## Testler

- Dönüşüm, ortalama sepet, iade oranı, değişim ve pay; ziyaret ve sipariş verisi olmayan ürün.
- İşaret sınırları (tam %15 iade, tam %30 düşüş), bu hafta satışı olmayan ürün, hafta seçimi.
- Kanal ve kategori toplamları, okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

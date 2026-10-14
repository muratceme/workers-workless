# Satış Hızı (Sell-Through) Raporu · Kod Bloğu

> Ürün Planlama ve Alokasyon › Alokasyon Uzmanı · Workers / Workless

Ürün ve mağaza bazında **satış oranını (sell-through)**, günlük satış hızını ve **stok gününü** hesaplar; indirime aday
ürünleri, hızlı satanları ve mağazalar arası **transfer** fırsatlarını işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Satış oranı % | Satılan ÷ (satılan + eldeki stok) × 100 |
| Günlük satış hızı | Son 7 gün satışı ÷ 7 (yazılmışsa); yoksa satılan ÷ girişten bu yana geçen gün |
| Stok günü | Eldeki stok ÷ günlük satış hızı |
| Sezon sonunda kalacak | Eldeki stok − günlük hız × sezon sonuna kalan gün |

Ürün değerleri mağazaların toplamıdır (ürün hızı = mağaza hızlarının toplamı).

| Öneri | Kural |
|---|---|
| İndirime aday | `--sezon-sonu` verilmişse stok günü sezon sonuna kalan günden fazla; verilmemişse satış oranı `--esik` (varsayılan %40) altında |
| Değerlendirmek için erken | İndirime aday kuralına uyuyor ama ürün `--min-hafta` (varsayılan 4) haftadan kısa süredir satışta |
| Hızlı satıyor | Stok günü `--az` (varsayılan 7) günden az |
| Transfer önerisi | Aynı üründe stok günü `--az` altında bir mağaza ile `--fazla` (varsayılan 60) üstünde bir mağaza var |

**Örnek:** Keten gömlek 3 mağazada 60 satılmış, 170 stokta → satış oranı **%26,09**. Günlük hız 0,66 → stok günü
**257,83**; sezon sonuna 46 gün var → **indirime aday**, sezon sonunda tahmini 140 adet kalır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                              # örnek veri (30.09.2026, sezon sonu 15.11.2026)
python main.py --satislar urun_magaza_satislari.xlsx --tarih 30.09.2026 --sezon-sonu 15.11.2026
python main.py --satislar urun_magaza_satislari.xlsx --esik 35 --min-hafta 3 --az 5 --fazla 45
```

Her satır bir ürünün bir mağazadaki durumudur. Zorunlu sütunlar: Ürün Kodu, Satış Adedi, Eldeki Stok. Mağaza sütunu
yoksa tek lokasyon sayılır. Gelen adet yazılırsa satış + stokla karşılaştırılır.

## Çıktı

`cikti/satis_hizi_raporu.xlsx`:

- `Ürünler`: öneriye ve satış oranına göre sıralı; satış oranı, hız, stok günü, sezon sonu tahmini, transfer önerisi.
- `Mağaza Detayı`: ürün × mağaza göstergeleri; hızlı satan ve satışı olmayan satırlar renkli.
- `Uyarılar`.

## Dikkat

- **Öneri karar değildir.** İndirim kararında marj, indirim takvimi, ürünün sezonluk mu sürekli mi olduğu ve
  bedenlerin dağılımı (kırık beden) gözetilmelidir. Kod beden ve renk kırılımı yapmaz; bunun için Ürün Kodu sütununa
  model-renk-beden yazabilirsiniz.
- **Hız doğrusal varsayılır.** Gelecek haftalarda satış hızının aynı kalacağı varsayılır; hava, kampanya ve mevsim
  geçişi hesaba katılmaz. Son 7 gün satışı yazmak hızı güncel tutar.
- **Satış oranı tanımı:** Gelen adet yerine satılan + eldeki stok kullanılır; transfer, iade ve kayıp bu ikisini
  farklılaştırır (uyarı verilir).
- **Transfer önerisi** taşıma maliyetini ve mağazaların beden ihtiyacını bilmez; yalnız aday çiftleri gösterir.
- **Eşikler politika varsayılanıdır.** %40, 4 hafta, 7 ve 60 gün genel kabul görmüş değerler değildir.
- **Örnek veri:** Ürünler, mağazalar ve sayılar kurgusaldır.

## Testler

- Satış oranı, son 7 gün ve girişten bu yana hız, stok günü, ürün toplamı.
- Eşik ve sezon sonu kuralı, erken ürün, satışı olmayan ürün, transfer önerisi, sezon sonunda kalacak adet.
- Okuma (yinelenen satır, gelen farkı) ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

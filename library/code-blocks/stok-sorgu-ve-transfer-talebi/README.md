# Stok Sorgu ve Transfer Talebi · Kod Bloğu

> Mağaza Operasyonları › Satış Danışmanı · Workers / Workless

Müşterinin istediği ürün, renk ve beden mağazada yoksa **diğer mağazaların stoklarını** listeler ve en uygun kaynak mağazadan
**transfer talebi** hazırlar. Kaynak mağazada son ürünü bırakır, yakın mağazaları önceler. İnternete ve mağaza yazılımına
bağlanmaz; dışa aktarılmış stok listesini okur.

## Nasıl çalışır?

| Adım | Kural |
|---|---|
| Eşleşme | Ürün kodu + renk + beden (büyük-küçük harf ve Türkçe karakter farkı önemsenmez) |
| Verilebilir adet | Stok − rezerve − `--birak` (varsayılan 1; kaynak mağazada en az bir ürün kalır) − aynı çalıştırmada başka taleplere ayrılan |
| Kaynak sırası | Aynı şehir → aynı bölge → verilebilir adedi çok olan → mağaza adı (mağaza listesi verilmezse yalnız adet ve ad) |
| Bölme | Talep birden fazla mağazadan karşılanabilir; `--tek-kaynak` ile yalnız tamamını karşılayabilen tek mağaza |

**Uyarılar:** okunamayan satır (Yüksek) · ürün / renk / beden stok listesinde yok · talep karşılanamadı ya da kısmen karşılandı · negatif
stok (Orta) · talep eden mağazada kullanılabilir stok zaten var (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (8 mağaza)
python main.py --stok stok.xlsx --talepler talepler.xlsx --magazalar magazalar.xlsx
python main.py --stok stok.xlsx --talepler talepler.xlsx --birak 0 --tek-kaynak
```

## Çıktı

- `cikti/transfer_talebi.xlsx`: `Transfer Talepleri`, `Stok Sorgusu` (her talep için stoku olan mağazalar: stok / verilebilir),
  `Karşılanamayan`, `Uyarılar`.
- `cikti/transfer_talebi_transfer.csv`: mağaza yazılımına ya da e-postaya aktarılacak transfer listesi.

## Dikkat

- **Stok listesi anlık değildir.** Transferden önce kaynak mağazayı arayıp ürünün fiziksel olarak bulunduğunu teyit edin.
- **Transfer kuralları şirketinize göre değişir** (bölgeler arası transfer yasağı, kargo maliyeti, outlet / ana mağaza ayrımı).
  `--birak` ve mağaza listesini buna göre ayarlayın.
- **Müşteri bilgisi** not sütununa yalnız gerektiği kadar yazılmalıdır (KVKK).
- **Örnek veri:** Mağazalar, ürünler ve stoklar kurgusaldır.

## Testler

- Şehir / bölge önceliği, bırakılacak adet, rezerve, aynı çalıştırmada ardışık talepler, kısmi karşılama, tek kaynak seçeneği.
- Uyarılar, satır birleştirme ve negatif stok, CSV çıktısı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

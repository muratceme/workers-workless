# Malzeme İhtiyaç Planlaması (MRP) · Kod Bloğu

> Üretim › Üretim Planlama › Üretim Planlama Uzmanı · Workers / Workless

Ana üretim planından (MPS) ve **çok seviyeli ürün ağacından** her malzeme için dönem bazında **standart MRP
kaydını** hesaplar. Hangi malzemenin, ne kadar, **hangi dönem sipariş edilmesi/üretime verilmesi** gerektiğini
söyler. İnternete bağlanmaz.

| Satır | Hesap |
|---|---|
| Brüt ihtiyaç | MPS (mamul) veya üst malzemenin planlanan verilişi × birim miktar × (1 + fire) |
| Planlanmış girişler | Açık satın alma/üretim siparişleri |
| Öngörülen eldeki stok | Önceki stok + giriş − brüt ihtiyaç (+ planlanan giriş) |
| Net ihtiyaç | Stok emniyet stoğunun altına düşerse aradaki fark |
| Planlanan sipariş girişi | Net ihtiyaç, lot kuralına göre yuvarlanmış hâli |
| Planlanan sipariş verilişi | Giriş − tedarik süresi |

- **Lot kuralları:** **L4L** (sipariş bazında) ve **Sabit** (lot miktarının katı); ardından **en az sipariş
  miktarı** ve **ambalaj katı** uygulanır.
- **Düşük seviye kodu:** Bir malzeme ağacın birden fazla seviyesinde geçebilir (ör. vida hem masada hem
  ayakta). Böyle bir malzeme tüm ihtiyaçlar toplandıktan sonra en alt seviyesinde bir kez hesaplanır. Ağaçtaki
  döngüler hata verir.
- **GEÇMİŞ / acil:** Veriliş dönemi planlama ufkunun başından önceye düşen siparişler işaretlenir. Hemen
  verilmeli, tedarik süresi kısaltılmalı ya da MPS kaydırılmalıdır.
- **Uyarılar:** Başlangıç stoğu emniyet stoğunun altında olan, malzeme listesinde bulunmayan ve açık sipariş
  termini geçmiş kalemler uyarılır.

> Sonsuz kapasite MRP'sidir. Üretim kapasitesini görmek için **Haftalık Üretim Çizelgesi** kullanın.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                          # örnek: 2 masa modeli, 10 malzeme, 8 hafta
python main.py --mps mps.xlsx --bom urun_agaci.xlsx --stok malzemeler.xlsx --acik acik_siparisler.xlsx --baslangic 02.11.2026 --donem 12
```

## Çıktı

`Sipariş Önerileri` · `MRP Kayıtları` (malzeme × dönem, ES altı kırmızı, verilişler sarı) · `Bilgi`

## Testler

Ders kitabı örneğiyle test edilir. A → 2B → 3C (tedarik süreleri 1-2-1, eldeki 20-50-100) için verilişler
A {4: 80, 7: 150}, B {2: 110, 5: 300}, C {1: 230, 4: 900} çıkar. Ayrıca şunlar doğrulanır: sabit lot, emniyet
stoğu, fire, en az sipariş, geçmiş sipariş, düşük seviye kodu ve döngü tespiti.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

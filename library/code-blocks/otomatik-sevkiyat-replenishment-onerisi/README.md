# Otomatik Sevkiyat (Replenishment) Önerisi · Kod Bloğu

> Perakende › Ürün Planlama ve Alokasyon › Alokasyon Uzmanı · Workers / Workless

Mağaza satış hızı, mağaza stoğu ve hedef stok gününe göre depodan mağazalara **sevk edilecek miktarları**
hesaplar. Depo sistemine aktarılabilir bir sevk listesi üretir. İnternete bağlanmaz.

## Hesap

```
günlük satış = son N günün satışı ÷ N                                (--satis-gunu, varsayılan 28)
hedef stok   = max(asgari teşhir, günlük satış × (hedef stok günü + sevk süresi))
ihtiyaç      = hedef stok − (mağaza stoğu + yoldaki)  → paket/koli katına yukarı yuvarlanır
```

**Depo yetersizse:** Ürünün toplam ihtiyacı depo stoğunu aşarsa her turda **stok günü en düşük** mağazaya bir
paket verilir. Böylece stoksuz kalmak üzere olan mağazalar önce beslenir ve kalan stok adil paylaşılır.

**Uyarılar**
- **Depo kaynaklı:** Depoda stok yok veya depo yetersiz.
- **Teşhir:** Sevkten sonra da asgari teşhirin altında kalan mağazalar.
- **Fazla stok:** Son dönemde hiç satmayan veya stok günü `--fazla-gun` (varsayılan 60) değerini aşan kalemler.
  Bunlar mağazalar arası **transfer adayıdır**.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                          # örnek: 5 mağaza, 6 ürün
python main.py --magaza-stok magaza_stok.xlsx --satislar satislar.xlsx --depo-stok depo.xlsx --hedef-gun 14
python main.py ... --satis-gunu 21 --sevk-suresi 1 --asgari 3 --paket paketler.xlsx --bitis 14.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Mağaza stoğu | Mağaza, Ürün Kodu, Stok, [Yoldaki, Asgari Teşhir, Ürün Adı] |
| Satışlar | Tarih, Mağaza, Ürün Kodu, Adet |
| Depo stoğu | Ürün Kodu, Stok (kullanılabilir — rezerv ve sipariş ayrılmış stok hariç) |
| Paketler | Ürün Kodu, Paket (sevk katı) |

## Çıktı

`Sevk Listesi` (mağaza, ürün, adet — depo sistemine aktarım için) · `Hesap Detayı` (satış hızı, stok günü, hedef,
ihtiyaç, öneri, sevk, notlar; en riskli kalemler üstte) · `Ürün Özeti` · `Bilgi`

## Dikkat

- **Satış hızı:** Stoksuz geçen günlerde satış olmadığı için satış hızı düşük görünür.
- **Özel dönemler:** Kampanya, yeni ürün ve sezon geçişlerinde geçmiş satış hızı yanıltıcı olabilir; bu kalemleri
  elle gözden geçirin.
- **Hedef stok günü:** Mağaza kapasitesine ve sevk sıklığına göre ayarlanmalıdır. Örneğin haftada iki sevk alan
  mağaza için daha düşük bir değer yeterlidir.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

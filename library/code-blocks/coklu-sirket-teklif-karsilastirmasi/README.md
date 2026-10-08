# Çoklu Şirket Teklif Karşılaştırması · Kod Bloğu

> Sigorta › Teknik ve Operasyon › Sigorta Uzmanı · Workers / Workless

Farklı sigorta şirketlerinden alınan teklifleri teminat, limit, muafiyet ve prim bazında **müşteriye sunulacak tek
bir tabloya** döker. Müşterinin talebini karşılamayan noktaları işaretler. İnternete bağlanmaz.

## Ne yapar?

- **Teminat eşleştirme:** Şirketler teminatları farklı yazsa da ("Yangın (bina)" / "Yangın bina") teminatlar
  müşterinin talep listesiyle eşleştirilir. Talepte olmayanlar "ek teminat" olarak gösterilir.
- **Limit kontrolü:** İstenen limitin altında kalan teminatlar işaretlenir.
- **Muafiyet kontrolü:** Farklı biçimlerdeki muafiyetler istenen limit tutarındaki bir hasar üzerinden TL'ye
  çevrilerek karşılaştırılır. Örnek: 11.000.000 TL'de %5 = 550.000 TL, azami %2 = 220.000 TL.
  - Tanınan biçimler: `%2`, `%2, en az 10.000 TL`, `%5 en fazla 1.000`, `5.000 TL`, boş / "Yok".
  - "7 gün" gibi süre muafiyetleri gösterilir ama karşılaştırılmaz.
- **Uygunluk:** Zorunlu teminatların hepsi var, limit ve muafiyetleri uygun, geçerlilik süresi dolmamış ve brüt
  primi bilinen teklifler "Uygun" sayılır. Uygun teklifler brüt prime göre sıralanır.
- **Öne çıkanlar:** En düşük primli uygun teklif ve tüm talebi (isteğe bağlılar dahil) karşılayan en düşük primli
  teklif ayrıca belirtilir. Seçim müşterinindir.
- **Uyarılar:** Teminat notları (ör. koasürans), özel şartlar, ek teminatlar, süresi dolmuş teklifler, prim bilgisi
  olmayan şirketler.

Talep dosyası verilmezse teklifler yalnız yan yana konur ve prime göre sıralanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                       # örnek işyeri paket sigortası, 4 kurgusal şirket
python main.py --teklifler teklifler.xlsx --ozet teklif_ozeti.xlsx --talep talep.xlsx
python main.py --teklifler teklifler.xlsx --ozet teklif_ozeti.xlsx --tarih 09.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Teklifler | Şirket, Teminat, Limit, Muafiyet, Not (şirket başına teminat satırları) |
| Teklif özeti | Şirket, Net Prim, Brüt Prim, Ödeme Planı, Geçerlilik Tarihi, Özel Şart |
| Talep | Teminat, İstenen Limit, Azami Muafiyet, Zorunlu (Evet / Hayır) |

## Çıktı

- **Excel:** `teklif_karsilastirma.xlsx`:
  - `Karşılaştırma`: teminat × şirket (limit ve muafiyet), ek teminatlar, prim, ödeme planı, geçerlilik, sonuç.
    Renkler: yeşil uygun, kırmızı talebi karşılamıyor, gri isteğe bağlı teminat yok, sarı kontrol edin.
  - `Değerlendirme`: sıra, eksik ve uygun olmayan teminatlar, uyarılar; boş "Müşteri Tercihi" sütunu.
- **Markdown:** `teklif_karsilastirma.md`, müşteriye gönderilebilecek özet tablo.

## Dikkat

- **Bilgi amaçlı:** Karşılaştırma bilgi amaçlıdır. Poliçe genel ve özel şartları, klozlar, istisnalar ve
  şirketlerin yazılı teklifleri esastır. Aynı adlı teminatın kapsamı şirketten şirkete farklı olabilir.
- **Muafiyet karşılaştırması:** Tek bir hasar tutarı üzerinden yapılan bir yaklaşımdır. Gerçek hasarda muafiyetin
  etkisi hasar tutarına göre değişir.
- **Eşleştirme:** Kelime benzerliğine dayanır. Çok farklı adlandırılmış teminatları teklif dosyasında talep
  listesindeki adla yazın.
- **Örnek veri:** Örnek şirketler ve teklifler kurgusaldır.

## Testler

Şunlar test edilir:
- Örnek 4 teklifin sıralaması ve uygunluk sonuçları; ad eşleştirme ve ek teminat.
- Limit ve muafiyet uygunsuzlukları; süresi dolmuş teklif; muafiyet ayrıştırma.
- Talepsiz kullanım ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

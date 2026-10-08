# Ürün Maliyeti Hesaplama · Kod Bloğu

> Muhasebe › Maliyet Muhasebesi Uzmanı · Workers / Workless

Çok seviyeli reçete (ürün ağacı) ve rota üzerinden ürün **birim maliyetini** hesaplar. Maliyet üç bileşene ayrılır:
direkt ilk madde ve malzeme, direkt işçilik ve genel üretim giderleri. İnternete bağlanmaz.

| Bileşen | Hesap |
|---|---|
| **DİMM** (malzeme) | reçete miktarı × (1 + fire %) ÷ reçete çıktı miktarı × malzeme birim maliyeti |
| **Direkt işçilik** | (süre dk/adet + hazırlık dk ÷ parti miktarı) ÷ 60 × işçilik TL/saat × kişi sayısı |
| **GÜG** | birim süre (saat) × iş merkezi GÜG TL/saat (veya `--gug-orani` ile işçiliğin yüzdesi) |

**Hesaplama ayrıntıları**
- **Yarı mamuller:** Önce kendi malzeme, işçilik ve GÜG'leriyle hesaplanır. Üst ürüne bu üç bileşen ayrı ayrı
  aktarılır, böylece mamulün toplam işçilik payı doğru görünür.
- **Dövizli malzemeler:** Para Birimi sütununda USD veya EUR yazan malzemeler `--kur` ile verilen kurla TL'ye
  çevrilir.
- **Hazırlık süresi:** Parti miktarına bölünerek birime dağıtılır. Bu sayede küçük partilerin maliyet etkisi
  görünür.
- **Brüt kâr marjı:** Satış fiyatları verilirse brüt kâr ve marj hesaplanır. Zarar eden ürünler kırmızı işaretlenir.
- **Hata denetimi:** Reçete döngüleri, maliyeti olmayan malzemeler ve ücreti olmayan iş merkezleri uyarıyla
  raporlanır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                   # örnek mobilya verileriyle (sandalye, masa)
python main.py --recete recete.xlsx --malzeme malzeme.xlsx \
               --rota rota.xlsx --is-merkezi is_merkezi.xlsx --parti partiler.xlsx \
               --fiyat fiyatlar.xlsx --kur USD=41,20 EUR=48,10
python main.py --recete recete.xlsx --malzeme malzeme.xlsx      # yalnız malzeme maliyeti
```

| Dosya | Sütunlar |
|---|---|
| Reçete | Ana Ürün, Bileşen, Miktar, [Birim, Fire %, Çıktı Miktarı] |
| Malzeme | Malzeme Kodu, [Malzeme Adı, Birim], Birim Maliyet, [Para Birimi] |
| Rota | Ürün Kodu, [Operasyon], İş Merkezi, Süre (dk/adet), [Hazırlık (dk), Kişi Sayısı] |
| İş merkezi | İş Merkezi, İşçilik (TL/saat), [GÜG (TL/saat)] |
| Parti | Ürün Kodu, Parti Miktarı |
| Fiyat | Ürün Kodu, Satış Fiyatı, [Para Birimi] |

**Değerlerin belirlenmesi**
- **Çıktı miktarı:** "Bu reçete kaç adet için yazıldı" bilgisidir. Örneğin 10 kg hammaddeden 5 adet çıkıyorsa
  Çıktı Miktarı 5 yazılır.
- **İşçilik saat ücreti:** Bordro maliyeti (brüt ücret + işveren primleri) ÷ fiilî çalışma saati olarak hesaplanır.
- **GÜG saat ücreti:** Dönemin GÜG bütçesi ÷ normal kapasite saati olarak hesaplanır.

## Çıktı

`Birim Maliyetler` (bileşen grafiği dahil) · `Maliyet Ağacı` · `Operasyonlar` · `Malzeme Payları` (hangi malzeme
maliyetin ne kadarı) · `Bilgi`

## Dikkat

- **Maliyetin türü:** Hesaplanan değer standart (ön) maliyettir. Gerçekleşen maliyetle aradaki fiyat, miktar ve
  verimlilik farkları ayrıca analiz edilmelidir.
- **Stok değerlemesi:** Yasal stok değerlemesi ve sabit giderlerin dağıtımı için mali müşavirinize danışın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

# Teklif Karşılaştırma · Kod Bloğu

> Satın Alma › Satın Alma Uzmanı · Workers / Workless

Tedarikçilerden gelen teklifleri tek tabloda toplar. **Döviz, nakliye ve ödeme vadesi farklarını
eşitleyerek** karşılaştırır, tedarikçileri ağırlıklı kriterlerle puanlar ve bir karşılaştırma tablosu (mukayese
tablosu) üretir. İnternete bağlanmaz.

## Neler yapar?

- **Ortak para birimi:** Dövizli teklifler `--kur` ile verdiğiniz kurla TL'ye çevrilir. Nakliye ve ek masraflar
  toplama eklenir.
- **Vade farkı:** Peşin ile 90 gün vadeli teklif aynı değildir. Tutarlar `--faiz` ile verilen yıllık iskonto
  oranıyla bugünkü değere indirgenir: `tutar ÷ (1 + oran × gün / 365)`.
- **Kalem bazında karşılaştırma:** Her malzeme için en uygun teklif, ikinciye göre fark ve **bölünmüş alım**
  toplamı (her kalemi en uygun tedarikçiden almak) hesaplanır.
- **Ağırlıklı puanlama:** Kriterler fiyat, termin, vade, kalite ve garantidir; varsayılan ağırlıklar 50 / 15 / 15
  / 15 / 5. Her kriterde en iyi teklif 100 puan alır.
  - Fiyat ve terminde puan = en iyi ÷ teklif × 100.
  - Vade, kalite ve garantide puan = teklif ÷ en iyi × 100.
  - Eksik kalemli tedarikçiler fiyatta yalnız teklif verdikleri kalemlerle kıyaslanır.
- **Eleme ve uyarılar:**
  - Teknik olarak uygun olmayan teklifler (`Teknik Uygun = H`) değerlendirmeye alınmaz.
  - Eksik kalemli teklifler işaretlenir; tek tedarikçi önerisi yalnız tüm kalemlere teklif verenlerden seçilir.
  - Medyanın %30'undan fazla altında kalan fiyatlar **aşırı düşük teklif** olarak işaretlenir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                    # örnek tekliflerle (örnek kur ve %40 faiz)
python main.py --girdi teklifler.xlsx --kur USD=41,20 EUR=48,05 --faiz 40
python main.py --girdi teklifler.xlsx --agirlik fiyat=60 termin=10 vade=10 kalite=15 garanti=5
```

**Girdi:** Her satır bir tedarikçinin bir kaleme verdiği tekliftir. Sütunlar: Tedarikçi, Malzeme Kodu,
Malzeme, Miktar, Birim, Birim Fiyat, Para Birimi, Nakliye, Termin (gün), Ödeme Vadesi (gün), Garanti (ay),
Kalite Puanı, Teknik Uygun (E/H), Not.

## Çıktı

`Değerlendirme` · `Kalem Bazında` · `Teklif Listesi` · `Bilgi`

> Puanlama bir karar destek aracıdır. Nihai seçim şartname uyumu, numune ve referanslarla birlikte yapılmalıdır.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

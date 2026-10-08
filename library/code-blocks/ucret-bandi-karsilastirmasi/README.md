# Ücret Bandı Karşılaştırması · Kod Bloğu

> İnsan Kaynakları › Ücretlendirme ve Yan Haklar Uzmanı · Workers / Workless

İç ücret verisini kademe veya pozisyon bazında ücret araştırması bantlarıyla kıyaslar. Bant dışı çalışanları,
asgari ücret altını ve ücret sıkışmasını işaretler. Bant altındakileri banda çekmenin maliyetini hesaplar.
İnternete bağlanmaz.

## Nasıl hesaplar?

- **Bant seçimi:** Önce pozisyona özel bant aranır (ör. Yazılım Geliştirici). Yoksa çalışanın kademe bandı
  kullanılır. Bantta Alt / Orta / Üst değerleri bulunur; araştırmanın P25 / P50 / P75 değerleri gibi.
- **Kısmi süreli çalışan:** Ücret, tam zamanlı karşılığına çevrilir (ücret ÷ çalışma oranı).
- **Karşılaştırma oranı (compa-ratio):** ücret ÷ orta nokta. **Bant içi konum:** (ücret − alt) ÷ (üst − alt).

| Kontrol | Önem |
|---|---|
| Tam zamanlı karşılık yasal asgari ücretin altında (yıl bazında, `tr_parametreler.json`) | Yüksek |
| Bant altı | Yüksek |
| Bant üstü | Orta |
| Ücret sıkışması: aynı pozisyonda en az 2 yıl daha kıdemli ve performansı eşit / daha iyi çalışan, yeni gelenin ücretinin %95'inden az alıyor (`--sikisma`) | Orta |
| Cinsiyete göre bant ortalama karşılaştırma oranı farkı %5'in üstünde (her grupta en az 2 kişi) | Bilgi |
| Bant tanımsız | Bilgi |

**Bütçe etkisi:** Bant altındakileri alt sınıra çekmenin aylık ve yıllık brüt maliyeti hesaplanır. İşveren SGK
payı bu maliyete dahil değildir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 28 çalışan, 5 kademe + 1 pozisyon bandı
python main.py --calisanlar calisanlar.xlsx --bantlar bantlar.xlsx
python main.py --calisanlar c.xlsx --bantlar b.xlsx --yil 2026 --sikisma 90
```

| Dosya | Sütunlar |
|---|---|
| Çalışanlar | Sicil No, Ad Soyad, Departman, Pozisyon, Kademe, Brüt Ücret (TL), İşe Giriş, Cinsiyet, Performans (A/B/C veya 1–5), Çalışma Oranı (1 veya 0,75) |
| Bantlar | Kademe, Pozisyon (boş = kademe bandı), Alt, Orta, Üst, Kaynak |

## Çıktı

`ucret_bandi_karsilastirmasi.xlsx`:
- `Çalışanlar`: bant, karşılaştırma oranı, bant içi konum, durum (renkli) ve işaretler; boş "Öneri / Karar" sütunu.
- `Uyarılar`.
- `Bütçe Etkisi`: kişi bazında aylık ve yıllık ek brüt maliyet.
- `Kademe Özeti`:
  - bant bazında medyan ücret, karşılaştırma oranı aralığı, bant dışı sayıları
  - departman özeti
  - kullanılan asgari ücret
- `Cinsiyet Karşılaştırması`.

## Dikkat

- **Bantlar:** Ücret araştırması bantları, katıldığınız araştırmanın veya kurum ücret politikanızın değerleridir.
  Örnek değerler kurgusaldır. Bantların aylık brüt olduğundan ve aynı döneme ait olduğundan emin olun.
- **Asgari ücret:** Değer, ortak `tr_parametreler.json` dosyasından yıl bazında okunur. Bu dosya resmî kaynak ve
  tarihle güncellenir.
- **Cinsiyet karşılaştırması:** Bir ön tarama göstergesidir. Bulunan fark tek başına ayrımcılık anlamına gelmez.
  Kıdem, performans ve iş içeriğiyle birlikte değerlendirilmelidir: 4857 sayılı Kanun md. 5, eşit veya eşit
  değerde iş için cinsiyet nedeniyle daha düşük ücret kararlaştırılamaz.
- **Kişisel veri:** Ücret verisi gizlidir. Raporu yetkili kişilerle sınırlı paylaşın.

## Testler

Şunlar test edilir:
- Bant seçimi: pozisyon önceliği.
- Kısmi süreli çalışanın tam zamanlı karşılığı; asgari ücret kontrolü.
- Bant altı ve üstü; ücret sıkışması; bütçe etkisi; cinsiyet karşılaştırması.
- Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

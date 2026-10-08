# Reklam Kampanyası Performans Raporu · Kod Bloğu

> Pazarlama › Dijital Pazarlama Uzmanı · Workers / Workless

Reklam platformlarından alınan dışa aktarımları **tek tabloda birleştirir** ve kampanya bazında performans
raporu üretir. Desteklenen platformlar Google Ads, Meta Ads, TikTok Ads ve Microsoft Ads'tir; Türkçe ve İngilizce
arayüz dökümleri okunur. İnternete bağlanmaz; platformlara API ile bağlanmaz.

## Ne yapar?

**Sütun tanıma**
- **Otomatik:** Başlıklar otomatik tanınır. Örneğin "Maliyet", "Amount spent (TRY)", "Harcanan tutar" ve "Cost"
  harcama olarak; "Dönüşümler", "Purchases" ve "Satın almalar" dönüşüm olarak okunur.
- **Ek satırlar:** Üstteki rapor adı ve tarih aralığı satırları ile alttaki "Toplam" satırı atlanır.
- **Sayı biçimleri:** Türkçe (1.234,56) ve İngilizce (1,234.56) biçimleri karışık olsa da doğru okunur.
- **Elle eşleştirme:** Tanınmayan sütunlar `--esleme harcama='Spend (TRY)'` ile eşleştirilir. Bilgi sayfasında
  her dosyada hangi sütunun hangi alana eşlendiği listelenir.

**Ölçütler**

| Ölçüt | Tanım |
|---|---|
| CTR | tıklama ÷ gösterim |
| TBM (CPC) | harcama ÷ tıklama |
| BGBM (CPM) | harcama ÷ gösterim × 1000 |
| Dönüşüm oranı | dönüşüm ÷ tıklama |
| CPA | harcama ÷ dönüşüm |
| ROAS | dönüşüm değeri ÷ harcama |

Ölçütler platform, kampanya ve hafta bazında hesaplanır.

**Değerlendirme işaretleri**
- **Harcıyor ama dönüşüm yok:** Toplam harcamanın en az %2'sini alan kampanyalar.
- **ROAS hedefin çok altında:** Hedefin %70'inin altı.
- **Bütçe artırılabilir:** ROAS hedefin %130'unun üstünde ve harcama payı en az %5.
- **CPA yüksek:** CPA hedefin %130'unun üstünde.
- **Tıklama oranı çok düşük:** CTR %0,5'in altında.

**Bütçe temposu:** `--butce` dosyası verilirse harcanan tutar, bugüne kadar beklenen harcamayla karşılaştırılır
(dönem bütçesi × geçen gün ÷ toplam gün). %85–115 dışı işaretlenir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                            # örnek: Google (TR) + Meta (EN), Eylül 2026
python main.py --dosya google_ads.csv --dosya meta_ads.csv --hedef-roas 4 --hedef-cpa 150
python main.py --dosya tiktok.xlsx --platform TikTok --butce butceler.xlsx
```

Platformdan **kampanya × gün** kırılımlı rapor indirin. Gün kırılımı haftalık tablo ve bütçe temposu için
gereklidir.

## Çıktı

`Özet` (platform bazında, grafik) · `Kampanyalar` (ölçütler, harcama payı, değerlendirme) · `Haftalık` ·
`Bütçe Temposu` · `Bilgi` (tanımlar ve tanınan sütunlar)

## Dikkat

- **Dönüşüm tanımları:** Platformların dönüşüm tanımları ve ilişkilendirme pencereleri farklıdır; aynı satış
  birden fazla platformda sayılabilir. Toplam ROAS'ı kendi sipariş verinizle (UTM) doğrulayın.
- **Para birimi:** Dosyalarda farklı para birimi varsa uyarı verilir; tutarları çevirip öyle birleştirin.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

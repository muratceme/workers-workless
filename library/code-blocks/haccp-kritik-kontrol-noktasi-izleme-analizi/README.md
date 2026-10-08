# HACCP Kritik Kontrol Noktası İzleme Analizi · Kod Bloğu

> Gıda Üretimi › Kalite Güvence ve Gıda Güvenliği › Kalite Güvence Uzmanı · Workers / Workless

Kritik kontrol noktası (KKN) izleme kayıtlarını (sıcaklık, süre, pH, metal dedektör…) HACCP planınızdaki kritik
ve operasyonel limitlerle karşılaştırır. Sapmaları, düzeltici faaliyet eksiklerini, izleme boşluklarını ve kayıt
bütünlüğü sorunlarını listeler. İnternete bağlanmaz.

## Ne kontrol eder?

| Kontrol | Önem |
|---|---|
| Kritik limit sapması; düzeltici faaliyet (DF) kaydı yoksa ayrıca vurgulanır | Kritik |
| Planda tanımlı KKN için hiç kayıt yok | Yüksek |
| İzleme boşluğu: aynı gün iki kayıt arası izleme sıklığının 1,5 katından uzun | Yüksek |
| Operasyonel (hedef) limit dışı ama kritik limit içinde | Orta |
| Limite yaklaşan eğilim: aynı gün art arda 5 ölçüm sürekli limit yönünde (`--egilim`) | Orta |
| Tekrar eden aynı değer: aynı gün art arda 8 kez birebir aynı değer (`--tekrar`) | Orta |
| Okunamayan değer; planda olmayan KKN kodu | Orta |
| Ölçen veya doğrulayan boş | Bilgi |

**Değer türleri:**
- **Sayısal KKN:** Kritik alt ve / veya üst limiti vardır.
- **Kategorik KKN:** Kritik limiti boş bırakılır, ör. metal dedektör test kartı. "Kaldı", "Red", "Uygunsuz" kritik
  sapmadır; "Geçti", "Uygun" uygundur.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                       # örnek süt işletmesi, 2 günlük kayıt
python main.py --plan kkn_tanimlari.xlsx --kayitlar izleme_kayitlari.xlsx
python main.py --plan plan.xlsx --kayitlar kayit.xlsx --egilim 6 --tekrar 10
```

| Dosya | Sütunlar |
|---|---|
| HACCP planı | KKN Kodu, Adım, Parametre, Birim, Kritik Alt, Kritik Üst, Operasyonel Alt, Operasyonel Üst, İzleme Sıklığı (dk), Tanımlı Düzeltici Faaliyet |
| İzleme kayıtları | Tarih, Saat, KKN Kodu, Değer, Parti / Lot, Ölçen, Düzeltici Faaliyet, Doğrulayan |

## Çıktı

`haccp_kkn_analizi.xlsx`:
- `Özet`: KKN bazında kayıt, kritik sapma, DF kaydı olmayan sapma, operasyonel limit dışı, uyum oranı, en düşük /
  en yüksek değer.
- `Sapmalar`: değer, limit, lot, kaydedilen ve tanımlı DF; boş "Ürün Durumu / QA Değerlendirmesi" sütunu.
- `Uyarılar`: "İnceleme" sütunuyla.
- `Kayıtlar`: tüm kayıtlar, durum ve işaretler (filtrelenebilir).
- `Grafikler`: sayısal KKN'lerin limit çizgili grafikleri.
- `HACCP Planı`.

## Dikkat

- **Limitler planınızdan:** Kritik limitler, izleme sıklıkları ve düzeltici faaliyetler kendi HACCP planınızdan
  girilmelidir. Örnekteki değerler (pastörizasyon 72 °C, soğuk depo 4 °C gibi) yalnız gösterim içindir.
- **Kapsam:** Analiz kayıtları denetler; sapmanın ürün güvenliğine etkisi, ürünün akıbeti ve kök neden
  değerlendirmesi gıda güvenliği ekibinin sorumluluğundadır.
- **İzleme boşluğu:** Aynı gün içindeki kayıtlara bakılır. Üretim yapılmayan aralıklar (mola, temizlik) da boşluk
  görünebilir.
- **Örnek veri:** Örnek işletme ve kayıtlar kurgusaldır.

## Testler

Şunlar test edilir:
- Plan okuma; 4 kritik sapma ve DF kaydı olmayan sapma.
- Tüm uyarı türleri ve parametreler.
- Kayıt yok / planda olmayan KKN / okunamayan değer; Excel sayfaları ve grafikler.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

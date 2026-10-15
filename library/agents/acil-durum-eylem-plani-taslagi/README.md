# Acil Durum Eylem Planı Taslağı · AI Agent

> İş Sağlığı ve Güvenliği › İş Güvenliği Uzmanı · Workers / Workless

İşyeri bilgilerinden **acil durum planı taslağı** hazırlar. Senaryoları, önleyici tedbirleri, ekiplere göre müdahale ve
tahliye adımlarını model yazar. **Ekip sayılarını, geçerlilik ve tatbikat tarihlerini kod hesaplar.** Risk değerlendirmesindeki
her risk planda ele alınmadıysa taslak modele geri gönderilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** İşyeri dosyasındaki bölümler B1, B2 …, riskler R1, R2 …, ekipmanlar E1, E2 … diye numaralanır.
2. **Hesap (kod):** İşyerlerinde Acil Durumlar Hakkında Yönetmelik (2021 değişikliğiyle) ve İlkyardım Yönetmeliği'ne göre:

   | Tehlike sınıfı | Söndürme, kurtarma, koruma (her biri) | İlk yardım | Plan yenileme |
   |---|---|---|---|
   | Çok tehlikeli | her 30 çalışana kadar 1 | her 10 çalışana kadar 1 | en geç 2 yıl |
   | Tehlikeli | her 40 çalışana kadar 1 | her 15 çalışana kadar 1 | en geç 4 yıl |
   | Az tehlikeli | her 50 çalışana kadar 1 | her 20 çalışana kadar 1 | en geç 6 yıl |

   10'dan az çalışanı olan işyerlerinde söndürme, kurtarma ve koruma için tek destek elemanı yeterlidir. Sayılar vardiya
   bazında hesaplanır (her vardiyada ekip bulunmalı). Tatbikat en geç yılda bir yapılır.
3. **Taslak (model):** Senaryolar (her biri dayandığı R/B/E numaralarıyla), önleyici tedbirler, ekiplere göre müdahale
   adımları ve tahliye adımları.
4. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bir risk (R) hiçbir senaryoda ele alınmamış ya da yangın / doğal afet senaryosu yok
   - Yangın senaryosunda söndürme, deprem senaryosunda kurtarma ekibine görev verilmemiş
   - Tahliyede toplanma yeri ya da sayım yok; gebe / engelli çalışan varken refakat, ziyaretçi varken yönlendirme yok
   - İşyeri bilgisinde olmayan sayı (ör. "5 kişilik ekip") ya da 110, 155, 156 gibi eski acil numaraları yazılmış
5. **Plan (kod):** Unvan, adres, hazırlayanlar, tarihler, ekip tablosu ve 112 irtibat bilgisi eklenir. Yönetmelik
   md. 12'de sayılan kroki unsurlarından (söndürücü, ilk yardım malzemesi, elektrik ve gaz kesim noktaları, alarm)
   işyeri dosyasında eksik olanlar ayrıca listelenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                               # örnek (kurgusal) plastik ambalaj fabrikası
python agent.py --girdi isyeri.docx --tur 3
```

İşyeri dosyası örneği:

```
Unvan: …
Tehlike Sınıfı: Tehlikeli
Vardiyalar: Gündüz 85; Gece 32
Toplanma Yeri: Ana giriş karşısındaki açık otopark
Özel Durumlar: Gebe çalışanlar var; ziyaretçiler sahada bulunuyor

Riskler:
- Granül stoklarında yangın
- Deprem nedeniyle raf devrilmesi

Ekipman:
- Üretim holünde 8 adet 6 kg KKT yangın söndürücü
- Elektrik ana kesme şalteri: trafo odası girişi
```

## Çıktı

- `cikti/acil_durum_plani.txt`: plan taslağı.
- `cikti/acil_durum_plani.xlsx`: Senaryolar; Ekipler (vardiya bazında en az destek elemanı, isim sütunu boş);
  Doküman Kontrol (md. 12 unsurları).

## Dikkat

- **Taslaktır.** Plan; risk değerlendirmesi, işyeri ziyareti ve çalışan katılımıyla iş güvenliği uzmanı, işyeri
  hekimi ve işveren tarafından tamamlanır, işverence onaylanır. Sayfalar paraflanır, son sayfa imzalanır.
- **Kroki ve ekip listesi** bu taslakta yoktur; tahliye planı giriş-çıkışlara ve katlara, ekip listesi görünür yere asılır.
- **Ek mevzuat:** Binaların Yangından Korunması Hakkında Yönetmelik, maden işyerleri ve büyük endüstriyel kaza
  kapsamındaki tesisler için ayrı hükümler saklıdır. Mevzuatın güncel hâlini kontrol edin.
- **Örnek veri:** İşyeri kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta üç riski atlar, yangında söndürme ekibine görev vermez, "4 kişilik
ekip" ve 110 numarası yazar, tahliyede sayım, refakat ve ziyaretçi yönlendirmesini unutur. Kodun dokuz sorunu da
yakaladığı, ekip sayılarını (gündüz 85 çalışan, tehlikeli: 3 / 3 / 3 / 6) ve tarihleri doğru hesapladığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve planın mevzuata uygunluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.

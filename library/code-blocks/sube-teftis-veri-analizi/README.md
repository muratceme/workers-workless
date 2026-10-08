# Şube Teftiş Veri Analizi · Kod Bloğu

> Bankacılık › Teftiş Kurulu › Müfettiş Yardımcısı · Workers / Workless

Şube işlem verisini personel listesiyle birlikte tarar; teftişte incelenecek **riskli örüntüleri** bulgu listesi,
personel bazında risk puanı ve işaretli işlem listesi olarak verir. İnternete bağlanmaz.

Bulgular örnek seçimi ve inceleme içindir; **tek başına usulsüzlük göstermez.** Her bulguyu işlem belgeleri ve
ilgili personelle görüşerek değerlendirin.

## Ne kontrol eder?

| Kural | Önem |
|---|---|
| Personelin kendi müşteri numarasında işlem | Yüksek |
| Limit aşımı: tutar personelin işlem limitini aşıyor, onaylayan yok | Yüksek |
| Kendi işlemini onaylama (dört göz ilkesi) | Yüksek |
| Personel yakınının hesabında işlem | Orta |
| Onaylayanın limiti de tutara yetmiyor | Orta |
| Mesai dışı işlem (varsayılan 08:30–17:30) | Orta |
| Hafta sonu veya resmî tatil günü işlem | Orta |
| Sık iptal: en az 3 iptal, oran ≥ %5 ve şube oranının ≥ 2 katı | Orta |
| İptal sonrası farklı tutarla yeniden işlem: 60 dk içinde aynı personel, müşteri ve işlem türü | Orta |
| Bölünmüş nakit işlem: aynı müşteri, aynı gün, her biri eşiğin altında ama toplamı eşiği aşan nakit işlemler | Orta |
| Personel-müşteri yoğunlaşması: personelin işlemlerinin ≥ %10'u ve en az 5 işlem tek müşteride | Bilgi |
| Başka bir personelin veya yakınının hesabında işlem; personel listesinde olmayan sicil | Bilgi |

**Risk puanı:** Personel Özeti'nde her bulgu Yüksek 3, Orta 2, Bilgi 1 puan sayılır; 5 ve üzeri renklendirilir.

**Tatiller:** Resmî tatiller ve 2025–2027 dini bayramları (Diyanet takvimi) dahildir. Arife yarım günleri
işlenmez; idari izin ve ek tatilleri `--tatiller` ile verin.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek şube verisi (Eylül 2026)
python main.py --islemler islemler.xlsx --personel personel.xlsx
python main.py --islemler islemler.xlsx --personel personel.xlsx --mesai 09:00-17:00 --cumartesi --bolunmus-esik 75000 --tatiller tatiller.csv
```

| Dosya | Sütunlar |
|---|---|
| İşlemler | İşlem No, Tarih, Saat, Şube, Personel (sicil), Müşteri No, İşlem Türü, Tutar, Durum (Tamamlandı / İptal), Onaylayan (sicil), Açıklama |
| Personel | Sicil, Ad Soyad, Görev, Şube, İşlem Limiti, Müşteri No (personelin kendi), Yakın Müşteri No (birden çoksa `\|` ile) |

- **Sütun adları:** Yaygın varyasyonlar tanınır (ör. "İşlemi Yapan", "Onay Veren"). Üstteki başlık satırları
  atlanır. Saat ayrı sütunda değilse tarih hücresindeki saat okunur.
- **İşlem türleri:** Adında "nakit" veya "vezne" geçen işlemler nakit sayılır.
- **Durum:** "İptal", "Ters kayıt" veya "Red" içeren işlemler iptal sayılır.
- **Personel listesi olmadan:** Kural çalışır ama kendi hesabı, yakın hesabı ve limit kontrolleri yapılamaz.

## Çıktı

`sube_teftis_analizi.xlsx`:
- `Özet`: dönem, işlem ve bulgu sayıları, bulgu türü dağılımı.
- `Bulgular`: önem, tür, personel, müşteri, ilgili işlemler; boş "Müfettiş Değerlendirmesi" sütunu.
- `Personel Özeti`: işlem, iptal, tutar, risk puanı, bulgu türlerine göre sayılar.
- `İşlemler`: tüm işlemler, "İşaretler" sütunuyla (filtrelenebilir).
- `Parametreler`: kullanılan eşikler.

## Dikkat

- **Eşikler kurumunuza göre:** Mesai saatleri, bölünmüş işlem eşiği (varsayılan 100.000 TL örnek bir iç eşiktir),
  iptal ve yoğunlaşma eşikleri kuruma özeldir; teftiş yönergenize göre ayarlayın. Mevzuattaki kimlik tespiti ve
  bildirim eşiklerini güncel MASAK düzenlemelerinden kontrol edin.
- **Personel kuralları:** Personelin kendi veya yakınının hesabında işlem yapmasına ilişkin kurallar bankanızın iç
  düzenlemelerine göre değerlendirilmelidir.
- **Kişisel veri:** İşlem ve personel verisi kişisel veri ve bankacılık sırrıdır; çıktıyı yalnız yetkili teftiş
  personeliyle paylaşın.
- **Örnek veri:** Örnek şube, personel ve müşteriler kurgusaldır.

## Testler

Örnek veride bilerek bırakılan 12 bulgunun tamamı, parametre değişiklikleri ve tatil, onaylayan limiti, tanımsız
personel kuralları test edilir.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

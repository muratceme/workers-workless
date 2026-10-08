# Fazla Mesai Hesaplama · Kod Bloğu

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Workers / Workless

Günlük giriş-çıkış kayıtlarından (PDKS dökümü) **4857 sayılı İş Kanunu**'na göre fazla çalışma, fazla sürelerle
çalışma ve genel tatil ücretini hesaplar. Yasal sınırları kontrol eder. İnternete bağlanmaz.

## Kurallar

| Konu | Hesap | Dayanak |
|---|---|---|
| Saat ücreti | aylık brüt ÷ 225 (30 gün × 7,5 saat) | uygulama |
| Fazla çalışma | haftalık (Pzt–Pz) 45 saati aşan süre × saat ücreti × **1,5** | md. 41 |
| Fazla sürelerle çalışma | sözleşmedeki süre 45'in altındaysa, o süre ile 45 arası × **1,25** | md. 41 |
| Denkleştirme | `--denklestirme N`: N haftalık dönemde ortalama 45 saat | md. 63 |
| Genel tatilde çalışma | çalışılan her gün için ayrıca bir günlük ücret (aylık ÷ 30); arifede 13.00 sonrası yarım gün | md. 47 |
| Ara dinlenmesi | mola yazılmamışsa asgari süre düşülür: ≤ 4 sa 15 dk, ≤ 7,5 sa 30 dk, daha uzun 1 sa | md. 68 |
| Serbest zaman | ücret yerine: fazla çalışmanın her saati için 1 sa 30 dk, fazla sürelerle çalışma için 1 sa 15 dk | md. 41 |
| Sınırlar | günlük 11 saat, gece çalışması 7,5 saat, yıllık 270 saat fazla çalışma, 7 gün üst üste çalışma | md. 63, 69, 41, 46 |

**Hesaplama ayrıntıları**
- **Gece vardiyası:** Gece yarısını geçen vardiyalar (22.00–07.00) desteklenir. Vardiya, başladığı güne yazılır.
- **Gece çalışması kontrolü:** Bir vardiyanın yarısından fazlası 20.00–06.00 arasına düşüyorsa gece çalışması
  sayılır ve 7,5 saat sınırı kontrol edilir.
- **Genel tatiller:** Puantaj Kontrolü paketindeki takvimden alınır. Takvim sabit tatilleri ve 2025–2027 dini
  bayramlarını içerir.

> **Tartışmalı nokta:** Genel tatilde çalışılan saatlerin haftalık 45 saat hesabına dahil edilip edilmeyeceği
> uygulamada tartışmalıdır (aynı saatin iki kez ücretlendirilmesi). Araç varsayılan olarak bu saatleri **dahil
> eder**. `--ubgt-haric` ile çıkarabilirsiniz; işyeri uygulamanızı hukuk danışmanınızla belirleyin.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                  # örnek: 3 personel, 3 hafta (19 Mayıs dahil)
python main.py --kayitlar pdks.xlsx --personel personel.xlsx
python main.py --kayitlar pdks.xlsx --personel personel.xlsx --denklestirme 8 --ubgt-haric
```

| Dosya | Sütunlar |
|---|---|
| Kayıtlar | Sicil, Ad Soyad, Tarih, Giriş, Çıkış, [Mola (dk)] — veya Sicil, Tarih, Çalışma Saati |
| Personel | Sicil, Aylık Brüt Ücret, [Saat Ücreti, Haftalık Sözleşme Saati, Yıl Başından FM] |

## Çıktı

`Personel Özeti` (saatler, brüt ek ödemeler, serbest zaman karşılığı, 270 saat kontrolü) · `Haftalık` ·
`Günlük Kayıtlar` (brüt süre, düşülen ara dinlenmesi, net çalışma) · `Bilgi` (kurallar ve sınır aşımları)

## Dikkat

- **Brüt tutarlar:** Hesaplanan tutarlar brüttür. SGK primi, gelir vergisi ve damga vergisi bordroda hesaplanır.
- **İşçi onayı:** Fazla çalışma için işçinin yazılı onayı gerekir (Fazla Çalışma Yönetmeliği).
- **Kayıtların doğruluğu:** PDKS kayıtlarındaki eksik veya hatalı okutmaları hesaplamadan önce düzeltin.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

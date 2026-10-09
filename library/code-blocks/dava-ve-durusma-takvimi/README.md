# Dava ve Duruşma Takvimi · Kod Bloğu

> Hukuk › Avukat · Workers / Workless

Dava listesinden **duruşma takvimini**, tebliğ tarihlerinden **kesin süre son günlerini** ve hatırlatma listesini
çıkarır. Takvim uygulamalarına (Outlook, Google Takvim, iPhone) aktarılabilen bir `.ics` dosyası da yazar. Her süre
için hesap adımları gösterilir. İnternete bağlanmaz.

## Nasıl hesaplar?

### Süre tablosu

Olay türü yazılır ve "Süre" sütunu boş bırakılırsa aşağıdaki kanuni süre kullanılır. Kesin süreler (HMK 94) ve
tabloda olmayan olaylar için süre "Süre" sütununa yazılmalıdır (ör. "2 hafta", "1 ay", "15 gün").

| Usul | Olay (metinde geçen) | Süre | Dayanak |
|---|---|---|---|
| HMK | cevap dilekçesi | 2 hafta | HMK 127 |
| HMK | cevaba cevap / ikinci cevap | 2 hafta | HMK 136 |
| HMK | bilirkişi raporuna itiraz | 2 hafta | HMK 281 |
| HMK | istinaf | 2 hafta | HMK 345 |
| HMK | istinafa cevap | 2 hafta | HMK 347 |
| HMK | temyiz | 1 ay | HMK 361 |
| İYUK | cevap (savunma, savunmaya cevap) | 30 gün | İYUK 16 |
| İYUK | istinaf | 30 gün | İYUK 45 |
| İYUK | temyiz | 30 gün | İYUK 46 |

Özel kanunlardaki süreler (ör. iş, tüketici, icra, ceza yargılaması) farklı olabilir. Bunlar için "Süre" sütununu
doldurun.

### Son gün

1. **Gün** olarak belirlenen sürede başlangıç günü sayılmaz. Tebliğ 18.09 + 30 gün = 18.10 olur.
2. **Hafta / ay** olarak belirlenen süre, son hafta / ayda başlangıç gününe karşılık gelen günde biter. O gün yoksa
   ayın son günü alınır (HMK 92). Tebliğ 31.08 + 1 ay = 30.09 olur.
3. **Adli tatil** (20 Temmuz – 31 Ağustos):
   - Adli tatile tabi işlerde son günü tatile rastlayan süre, tatilin bittiği günden itibaren bir hafta uzar ve
     7 Eylül'de biter (HMK 104; idari yargıda İYUK 8).
   - Son günü tatilden sonraya düşen süre uzamaz.
   - Dava listesinde "Adli Tatile Tabi = Hayır" yazılan işler (HMK 103) uzamaz.
4. **Tatil:** Son gün Cumartesi, Pazar veya resmî tatile rastlarsa izleyen ilk iş gününe kayar (HMK 93).
   - Ulusal bayram ve genel tatiller kodda tanımlıdır.
   - Dinî bayramlar her yıl değiştiği için tatil dosyasından okunur.
   - Son gün yarım güne (arife, 28 Ekim) denk gelirse uyarı verilir.
5. **İç hedef:** Kanuni son günden `--tampon` (3) iş günü önce. Rapor tarihinden önceye düşerse rapor tarihi
   gösterilir.

### Duruşmalar

- Rapor tarihinden `--gun` (60) gün sonrasına kadar olan duruşmalar takvime alınır.
- Hatırlatmalar duruşma ve son günden 7, 3 ve 1 iş günü önceye yazılır.

| Kontrol | Önem |
|---|---|
| Son günü geçmiş ve "Tamamlandı" işaretlenmemiş süre | Yüksek |
| Süre belirlenemedi (tabloda yok, "Süre" sütunu boş) | Yüksek |
| Son güne 3 gün veya daha az kaldı | Yüksek |
| Son güne 4–7 gün kaldı | Orta |
| Aynı avukatın aynı gün iki duruşması: 120 dakikadan yakınsa Yüksek, değilse Orta | Yüksek / Orta |
| Açık davada geçmiş duruşma tarihi (yeni tarih girilmemiş) | Orta |
| Tatil gününe duruşma; son gün yarım gün; tatil dosyasında o yıl yok; dava listesinde olmayan dosya | Orta |
| Adli tatil sınırı: süre tatil öncesi hafta sonunda bitiyor, izleyen iş günü adli tatilde | Orta |
| Adli tatile tabi işte tatil içinde duruşma; duruşma tarihi yok | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek: 12 dava, 12 süre olayı, rapor tarihi 09.10.2026
python main.py --davalar davalar.xlsx --sureler sureler.xlsx --tatiller tatiller.csv --bugun 09.10.2026 --gun 60 --tampon 3
```

| Dosya | Sütunlar |
|---|---|
| Davalar | Dosya No, Mahkeme, Usul (HMK / İYUK), Dava Türü, Müvekkil, Karşı Taraf, Müvekkil Sıfatı, Sorumlu Avukat, Duruşma Tarihi, Duruşma Saati, Adli Tatile Tabi (Evet / Hayır), Durum (Derdest / Karar / Kesinleşti...), Not |
| Süreler | Dosya No, Olay, Başlangıç Tarihi (tebliğ), Süre, Açıklama, Tamamlandı (Evet / Hayır) |
| Tatiller | Tarih, Açıklama, Yarım Gün (Evet) |

## Çıktı

- `dava_takvimi.xlsx`:
  - `Takvim`: duruşmalar ve son günler, tarih sırasıyla.
  - `Süreler`: kanuni son gün, gün adı, iç hedef, durum ve **hesap adımları**.
  - `Duruşmalar`, `Avukat Bazında` (açık dava ve 30 günlük yük), `Hatırlatmalar`, `Uyarılar`.
- `dava_takvimi.ics`: Her kayıt için 7 ve 1 gün önce hatırlatıcılı takvim etkinliği. Duruşmalar saatli,
  son günler tüm gün etkinliğidir.

## Dikkat

- **Süre hesabı yardımcıdır, sorumluluk avukattadır.** Tebliğ tarihini (e-tebligatta tebliğ edilmiş sayılma tarihi
  dahil) doğru girin. Kesin süreleri ara karardaki ifadeyle kontrol edin.
- **Mevzuat değişebilir.** Tablodaki süreler 7251 sayılı Kanun sonrası HMK ve İYUK metinlerine göredir. Kullanmadan önce
  güncel metni kontrol edin.
- **Adli tatil istisnaları** (HMK 103: ihtiyati tedbir, nafaka, kira tespiti ve tahliye gibi işler) dava bazında
  "Adli Tatile Tabi = Hayır" ile işaretlenmelidir.
- **Tatil dosyası** her yıl Resmî Gazete / Diyanet takvimine göre güncellenmelidir. Bayram tatili uzatmaları idari
  izindir ve adli süreleri uzatmaz; bunları tatil dosyasına yazmayın.
- **Örnek veri:** Örnek dosya numaraları, taraflar ve tarihler kurgusaldır.

## Testler

Şunlar test edilir:
- Gün / hafta / ay hesabı, ay sonu kuralı, hafta sonu ve bayram kayması.
- Adli tatil uzaması; tabi olmayan iş ve son günü tatilden sonraya düşen süre.
- Örnek veride süre tablosu (HMK / İYUK), iç hedef ve uyarılar.
- Takvim sırası, .ics yapısı (satır uzunluğu), Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Hukuki görüş yerine
geçmez. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

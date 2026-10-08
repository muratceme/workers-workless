# Mülakat Takvimi Planlama · Kod Bloğu

> İnsan Kaynakları › İnsan Kaynakları Uzman Yardımcısı · Workers / Workless

Aday ve mülakatçı müsaitliklerini eşleştirir; çakışmasız bir mülakat takvimi ve davet listesi çıkarır. Takvime
eklenebilen bir `.ics` dosyası da üretir. İnternete bağlanmaz.

## Nasıl planlar?

- **Panel:** Her pozisyonun gerekli rolleri (ör. İK + Teknik) aynı saatte panel olarak görüşür. Mülakat süresi ve
  formatı pozisyon tablosundan alınır.
- **Mülakatçı atama:**
  - Mülakatçı yalnız yetkili olduğu pozisyonlara atanır; "Tümü" yazılırsa her pozisyona girer.
  - Günlük en çok mülakat sayısı korunur.
  - İki mülakat arasında tampon süre bırakılır (varsayılan 15 dk, `--tampon`).
- **Sıralama:**
  - Müsaitliği en dar olan aday önce yerleştirilir.
  - Adayın uygun aralıklarında 15 dakikalık adımlarla en erken uygun saat seçilir.
  - Aynı roldeki mülakatçılardan en az mülakatı olan atanır.
- **Yerleşemeyen aday:** Neden yazılır. Ya bir rolde adayın zamanlarında hiç boş mülakatçı yoktur, ya da roller
  aynı anda boş değildir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 9 aday, 5 mülakatçı, 3 pozisyon
python main.py --adaylar adaylar.xlsx --mulakatcilar mulakatcilar.xlsx --pozisyonlar pozisyonlar.csv
python main.py --adaylar a.xlsx --mulakatcilar m.xlsx --pozisyonlar p.csv --tampon 10 --sirket "Örnek A.Ş."
```

| Dosya | Sütunlar |
|---|---|
| Adaylar | Aday No, Ad Soyad, Pozisyon, E-posta, Müsaitlik |
| Mülakatçılar | Mülakatçı, Rol, Pozisyonlar (virgülle; "Tümü"), Günlük En Çok, Müsaitlik |
| Pozisyonlar | Pozisyon, Gerekli Roller (virgülle), Süre (dk), Format |

Müsaitlik `13.10.2026 09:00-12:00; 14.10.2026 13:30-17:00` biçiminde yazılır. Müsaitlik son sütundaysa CSV'de
";" ile bölünmesi sorun olmaz; parçalar birleştirilir.

## Çıktı

- `mulakat_takvimi.xlsx`:
  - `Takvim`: tarih, saat, aday, pozisyon, format ve panel; boş "Davet Gönderildi" ve "Aday Onayı" sütunları.
  - `Mülakatçı Programı`: kişi bazında mülakatlar ve yük özeti (mülakat sayısı, dakika, doluluk).
  - `Yerleşemeyenler`: neden ve boş "Aksiyon" sütunu.
  - `Davet Metinleri`: adaya gönderilecek e-posta taslakları. Bağlantı / adres ve imza yer tutucuyla bırakılır.
- `mulakat_takvimi.ics`: Outlook / Google Takvim'e içe aktarılabilir etkinlikler (Europe/Istanbul saati).

## Dikkat

- **Plan bir öneridir:** Yerleştirme açgözlü bir yaklaşımla yapılır, en çok adayı yerleştiren çözümü garanti
  etmez. Yerleşemeyen adaylar için mülakatçılardan ek saat isteyin veya aday sırasını değiştirin.
- **Kapsam dışı:** Toplantı odası planlaması yapılmaz. Yüz yüze mülakatlar için oda ayrıca ayarlanmalıdır.
- **Davetler:** Davet metinleri taslaktır. Gönderimden önce yer tutucuları doldurun.
- **Kişisel veri:** Aday iletişim bilgileri KVKK kapsamında korunmalıdır. Davetleri BCC kullanmadan, kişiye özel
  gönderin.
- **Örnek veri:** Örnek adaylar ve mülakatçılar kurgusaldır.

## Testler

Şunlar test edilir:
- Müsaitlik ayrıştırma; panel ve rol yetkisi.
- Tampon süre ve günlük sınır; en dar müsaitlik önceliği.
- Yerleşemeyen nedenleri; .ics dosyası, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

# Numune Takip Çizelgesi · Kod Bloğu

> Tekstil ve Konfeksiyon › Müşteri Temsilciliği (Merchandising) › Development Merchandiser · Workers / Workless

Numune taleplerini takip eder: proto, fit, size set, PP (pre-production), SMS (salesman sample) ve TOP (top of
production). Her numunenin aşamasını, gönderim tarihini ve müşteri yanıtını izler. Geciken numuneleri, yanıt
beklenenleri ve açılmamış revizyonları aksiyon listesine çevirir. İnternete bağlanmaz.

## Ne kontrol eder?

| Kontrol | Önem |
|---|---|
| Planlanan kesim tarihine "kesimden önce onay" süresinden az kaldı (örnekte 7 gün) ve onaylı PP numunesi yok | Kritik |
| Gönderilmemiş numunenin termini geçti | Yüksek |
| Müşteri revize / red dedi, yeni revizyon açılmamış | Yüksek |
| Termine 3 gün veya daha az kaldı (`--uyari-gun`) | Orta |
| Müşteriye gönderildi, yanıt süresi aşıldı (hatırlatma listesi) | Orta |
| Fit numunesi onaylanmadan PP numunesi açılmış | Orta |
| Kayıt hatası: yanıt gönderimden önce, gönderim talepten önce, gönderim yokken yanıt var | Orta |
| Aynı model ve türde 3 veya daha fazla revize; yorumlu onaydaki yorumun aktarılması; AWB'siz gönderim; termin yok | Bilgi |

**Kurallar:**
- **Termin:** "İstenen Tarih" kullanılır. Bu tarih boşsa termin, talep tarihi + türün hazırlık süresidir.
- **Müşteri yanıtı:** "Onay", "Yorumlu onay" (approved with comments), "Revize" / "Red", "İptal" olarak okunur.
  İngilizce karşılıkları da tanınır.
- **Yanıt süresi:** Süre tablosunda yoksa 7 gündür (`--yanit-gun`). SMS için yanıt beklenmez.
- **Gün hesabı:** Takvim günüyle yapılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 18 numune, 6 model, durum tarihi 08.10.2026
python main.py --numuneler numuneler.xlsx --sureler numune_sureleri.csv
python main.py --numuneler numuneler.xlsx --sureler numune_sureleri.csv --siparisler siparisler.xlsx --bugun 15.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Numuneler | Numune No, Model, Müşteri, Sezon, Numune Türü, Revizyon, Talep Tarihi, İstenen Tarih, Gönderim Tarihi, Kargo / AWB No, Yanıt Tarihi, Müşteri Yanıtı, Müşteri Yorumu, Sorumlu |
| Numune süreleri | Numune Türü, Hazırlık Süresi (gün), Müşteri Yanıt Süresi (gün), Kesimden Önce Onay (gün) |
| Siparişler (isteğe bağlı) | Sipariş No, Model, Müşteri, Planlanan Kesim Tarihi, Sevk Tarihi |

Her revizyon ayrı bir satırdır (Fit R1, Fit R2…).

## Çıktı

`numune_takip.xlsx`:
- `Aksiyonlar`: önem sırasıyla, sorumlu ve model bazında yapılacaklar ile boş "Yapılan / Tarih" sütunu.
- `Model Durumu`: model × numune türü matrisi. Her hücre son revizyonu ve durumunu gösterir (ör. "R2 · Yorumlu
  onay 04.09").
- `Numuneler`: tüm satırlar; termin, termin kaynağı, durum, gecikme günü ve işaretler (filtrelenebilir).
- `Performans`: türe ve sorumluya göre zamanında gönderim oranı, müşteri bazında ortalama ve en uzun yanıt süresi.

## Dikkat

- **Süreler örnektir:** `numune_sureleri.csv` içindeki hazırlık, yanıt ve kesim öncesi onay süreleri örnektir. Kendi
  modelhane kapasitenize ve müşteri anlaşmanıza göre güncelleyin.
- **Sıra kuralı:** "Fit onayı olmadan PP" kontrolü, modelin fit numunesi kaydı varsa yapılır. Fit aşamasını atlayan
  müşterilerde uyarı çıkmaz.
- **Örnek veri:** Örnek müşteriler, modeller ve numuneler kurgusaldır.

## Testler

Şunlar test edilir:
- Tür ve yanıt eşleştirme; termin kaynağı.
- Kesim öncesi PP onayı; revize zinciri; gecikme ve yanıt bekleme.
- Fit onayı olmadan PP; veri hataları; performans; Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

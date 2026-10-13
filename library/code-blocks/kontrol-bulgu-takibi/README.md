# Kontrol Bulgu Takibi · Kod Bloğu

> İç Kontrol ve Risk Yönetimi › İç Kontrol Uzmanı · Workers / Workless

İç kontrol bulgularını **termin, birim ve önem derecesine göre** izler: gecikmiş, termini yaklaşan, termini birkaç kez
ertelenmiş ve tekrar eden bulguları öne çıkarır; kapanma oranlarını ve açık bulguların yaşını raporlar. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Kapalı bulgu | Kapanış tarihi dolu. Kapanış ≤ termin (ya da termin yok) ise "Zamanında kapandı", değilse "Geç kapandı" |
| Açık bulgu | Termin geçmişse "Gecikmiş"; termin yoksa "Termin yok"; termine `--gun` (15) gün veya az kaldıysa "Yaklaşıyor"; değilse "Açık" |
| Gecikme (gün) | Durum tarihi − termin (açık bulgu); kapanış − termin (geç kapanan bulgu) |
| Yaş (gün) | Durum tarihi − açılış tarihi; dilimler 0–30, 31–90, 91–180, 181+ |
| Kapanma oranı % | Kapalı ÷ toplam × 100 |
| Zamanında kapanma % | Zamanında kapanan ÷ kapalı × 100 |
| Ortalama kapanma süresi | Kapalı bulgularda (kapanış − açılış) ortalaması |
| Çok ertelenen bulgu | Termin revizyon sayısı `--revizyon` (2) ve üstünde olan açık bulgu |

Açık bulgular şu sırayla listelenir: gecikmiş → termini olmayan → yaklaşan → diğerleri; içinde önem derecesi ve
gecikme süresine göre.

**Örnek:** Termin 30.06.2026, durum tarihi 12.10.2026, bulgu hâlâ açık → **104 gün gecikmiş**. Önem derecesi Yüksek,
termin 2 kez ertelenmiş ve bulgu tekrar ediyor → raporun en üstünde yer alır ve üç ayrı uyarı üretir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --bulgular bulgular.xlsx
python main.py --bulgular bulgular.xlsx --gun 30 --revizyon 1 --tarih 31.10.2026
```

Zorunlu sütunlar: Bulgu No, Açılış Tarihi. Önem "Yüksek", "Orta" ya da "Düşük" yazılır; boş ya da başka bir değer
"Belirsiz" sayılır. Termin her ertelendiğinde revizyon sayısını bir artırın.

## Çıktı

`cikti/kontrol_bulgu_takibi.xlsx`:

- `Açık Bulgular`: öncelik sırasıyla; durum, önem, termin, gecikme, yaş, aksiyon sahibi, revizyon ve tekrar bilgisi.
- `Özet`: açık, gecikmiş ve kapalı bulgu sayıları, kapanma oranları, ortalama süreler, açık bulguların yaş dağılımı.
- `Birimler`, `Önem Dereceleri`: aynı göstergeler kırılımlı.
- `Tüm Bulgular` ve `Uyarılar`.

## Dikkat

- **Kod bulguyu değerlendirmez.** Önem derecesini, termini ve bulgunun kapandığını siz belirlersiniz; kod aksiyonun
  bulguyu gerçekten giderip gidermediğine, kapanış kanıtına ya da doğrulama testine bakmaz. Bir bulguyu kapatmadan
  önce kontrolün etkinliğini yeniden test etmek metodolojinizin gereğidir.
- **Zamanında kapanma oranı, revize edilmiş termine göredir.** Termini üç kez ertelenip son terminde kapanan bulgu
  "zamanında" görünür; bu yüzden revizyon sayısını ayrıca izleyin ve oranı tek başına başarı göstergesi saymayın.
- **Kapanma oranı dönemi dikkate almaz.** Dosyadaki bütün bulgular üzerinden hesaplanır; dönemsel oran için yalnız
  o dönemde açılan ya da kapanan bulguları ayrı dosyada çalıştırın.
- **Tekrar bilgisi sizin işaretinizdir.** Kod bulgu metinlerini karşılaştırıp benzerlik aramaz.
- **Kapsam dışı:** risk kabulü ve yönetim onayı süreci, aksiyon adımlarının ayrı ayrı izlenmesi, bulgu derecelendirme
  metodolojisi, üst yönetim ve denetim komitesi raporlaması.
- **Örnek veri:** Bulgular, birimler ve tarihler kurgusaldır.

## Testler

- Durum sınırları: termin dün / bugün / 15 ve 16 gün sonra, termini olmayan bulgu, zamanında ve geç kapanış.
- Yaş dilimleri, gecikme, kapanma süresi, oranlar ve ortalamalar; açık bulgu sıralaması.
- Revizyon eşiği, tekrar eden bulgu, önem derecesine göre uyarı düzeyi, birim ve önem kırılımı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

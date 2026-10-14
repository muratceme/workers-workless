# e-Defter Berat Yükleme Takibi · Kod Bloğu

> Muhasebe ve Defter İşleri › Muhasebe Elemanı (Büro) · Workers / Workless

Büronun takip ettiği mükelleflerin aylık e-defter dönemleri için **oluşturma, imzalama ve berat yükleme son tarihlerini**
hesaplar; yüklenmiş, yaklaşan ve süresi geçmiş dönemleri listeler. İnternete ve GİB / e-Defter sistemine bağlanmaz.

## Nasıl hesaplar?

| Kural | Son tarih |
|---|---|
| Aylık yükleme tercihi | Dönemi izleyen **4. ayın 10'u** (gelir vergisi mükellefi) / **14'ü** (kurumlar vergisi ve diğer mükellefler) |
| Hesap döneminin son ayı (aralık) | Yıllık beyannamenin verileceği ayı izleyen ayın 10'u (gelir vergisi: nisan) / 14'ü (kurumlar vergisi: mayıs) |
| Tatil kaydırma | Son gün hafta sonuna ya da resmî tatile denk gelirse izleyen ilk iş günü |
| Süre uzatımı | `--uzatma ESKİ=YENİ` (ör. `11.05.2026=30.06.2026`); zincirleme uzatmalar birlikte verilebilir |

Dayanak: Elektronik Defter Genel Tebliği (Sıra No: 1) 4.3.4 numaralı bölüm; 21.05.2024'te süreler 10 gün uzatıldı, 08.11.2024
değişikliğiyle gelir vergisi mükellefleri için 10, diğerleri için 14 gün kuralı geldi. Kontrol: 2024/Temmuz için kod
**11.11.2024** (10 Kasım pazar) ve **14.11.2024** üretir; bu, değişiklikle duyurulan tarihlerle aynıdır.

Takip edilen dönemler: seçilen yılın tamamlanmış ayları ile son tarihi bu yıla düşen önceki yıl dönemleri. Durumlar: Süresi
geçti · Yaklaşıyor (`--uyari` 10 gün içinde) · Bekliyor · Zamanında yüklendi · Geç yüklendi.

Resmî tatiller: sabit ulusal tatiller (1 Ocak, 23 Nisan, 1 Mayıs, 19 Mayıs, 15 Temmuz, 30 Ağustos, 29 Ekim) ve 2026 dini
bayramları (20–22 Mart, 27–30 Mayıs) kodda; diğer yılların dini bayramlarını ve idari izinleri `--tatiller` dosyasıyla ekleyin.
Arife yarım günleri iş günü sayılır.

**Uyarılar:** süresi geçmiş ve yüklenmemiş dönem (Yüksek) · son tarihi yaklaşan dönem · geçici vergi dönemi tercihli mükellef
(son tarihi hesaplanmaz) · okunamayan satır (Orta / Yüksek) · geç yüklenmiş dönem · imza tarihi yüklemeden sonra (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (14.10.2026)
python main.py --mukellefler mukellefler.xlsx --yuklemeler yuklemeler.xlsx
python main.py --mukellefler m.xlsx --yuklemeler y.xlsx --uzatma 11.05.2026=10.06.2026 --uzatma 10.06.2026=30.06.2026
python main.py --mukellefler m.xlsx --yuklemeler y.xlsx --tatiller tatiller.csv --yil 2027 --uyari 15
```

Dönem `2026/09`, `2026-9`, `09.2026` ya da `Eylül 2026` biçiminde yazılabilir.

## Çıktı

`cikti/e_defter_berat_takibi.xlsx`: `Dönem Takibi` (süresi geçen ve yaklaşanlar üstte), `Son Tarih Takvimi` (gün bazında
mükellef listesi), `Uyarılar`.

## Dikkat

- **Süreler sık değişir ve uzatılır.** GİB 2026'da da birkaç kez süre uzatımı duyurdu. Son tarihleri her ay edefter.gov.tr
  duyuruları ve güncel tebliğ metniyle karşılaştırın, uzatmaları `--uzatma` ile girin.
- **Geçici vergi dönemi tercihi, özel hesap dönemi ve tasfiye / işi bırakma** durumları bu sürümde hesaplanmaz.
- **Gecikmenin sonucu** (ceza, tasdik) mevzuata göre değerlendirilmelidir; rapor yalnız takip içindir.
- **Örnek veri:** Mükellefler, VKN'ler ve tarihler kurgusaldır.

## Testler

- Tebliğ örneği (2024/Temmuz), aralık dönemi, hafta sonu ve tatil kaydırma, ek tatil, zincirleme uzatma.
- Durum sınıflandırma, önceki yıl dönemleri, uyarılar; dönem ve tür tanıma, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

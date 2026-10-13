# DÖF Takibi · Kod Bloğu

> Kalite Güvence ve Kalite Kontrol › Kalite Mühendisi · Workers / Workless

Düzeltici ve önleyici faaliyetleri **sorumlu, termin ve etkinlik doğrulamasıyla** izler: gecikmiş aksiyonları,
tamamlandığı hâlde etkinliği doğrulanmamış faaliyetleri ve etkisiz bulunup yeniden açılması gerekenleri öne çıkarır.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Aşama | Durum | Kural |
|---|---|---|
| Aksiyon tamamlanmamış | Gecikmiş | Termin durum tarihinden önce |
| | Termin yok | Termin yazılmamış |
| | Termin yaklaşıyor | Termine `--gun` (15) gün veya az kaldı |
| | Açık | Diğerleri |
| Aksiyon tamamlanmış | Doğrulama bekliyor | Etkinlik doğrulama tarihi boş |
| | Doğrulama gecikti | Tamamlanmanın üzerinden `--dogrulama` (30) günden fazla geçmiş |
| Doğrulama yapılmış | Kapalı (etkin) | Etkinlik sonucu "Etkin" |
| | Etkisiz, yeniden açılmalı | Etkinlik sonucu "Etkisiz" |
| | Sonuç yazılmamış | Doğrulama tarihi var, sonuç boş |

| Gösterge | Formül |
|---|---|
| Kapanma oranı % | Kapalı (etkin) ÷ toplam × 100 |
| Etkinlik oranı % | Etkin bulunan ÷ (etkin + etkisiz) × 100 |
| Zamanında aksiyon % | Terminde ya da önce tamamlanan ÷ termini olan tamamlanmış faaliyet × 100 |
| Ortalama aksiyon süresi | Aksiyon tamamlanma − açılış (gün) |
| Ortalama kapanma süresi | Etkinlik doğrulama − açılış (gün), kapalı faaliyetlerde |

Bir DÖF ancak etkinliği doğrulanıp "Etkin" bulunduğunda kapalı sayılır; aksiyonun tamamlanması kapanış değildir.

**Örnek:** Aksiyon 10.09.2026'da tamamlandı, etkinlik doğrulaması yapılmadı; durum tarihi 12.10.2026 → **32 gündür**
doğrulama bekliyor, durum "Doğrulama gecikti".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek veri, durum tarihi 12.10.2026
python main.py --doflar dof_listesi.xlsx
python main.py --doflar dof_listesi.xlsx --gun 7 --dogrulama 60 --kritik 15 --tarih 31.10.2026
```

Zorunlu sütunlar: DÖF No, Açılış Tarihi. Etkinlik sonucu "Etkin" ya da "Etkisiz" yazılır ("Etkili", "Uygun",
"Evet" ve olumsuzları da tanınır).

## Çıktı

`cikti/dof_takibi.xlsx`:

- `DÖF Listesi`: en sorunludan başlayarak; durum, tarihler, gecikme ve doğrulama bekleme süresi.
- `Özet`: sayılar, kapanma, etkinlik ve zamanında aksiyon oranları, ortalama süreler, durum dağılımı.
- `Sorumlular`, `Kaynaklar`, `Türler`: aynı göstergeler kırılımlı.
- `Uyarılar`.

## Dikkat

- **Kod faaliyetin etkin olup olmadığına karar vermez.** Kök neden analizinin doğruluğunu, aksiyonun kök nedeni
  giderip gidermediğini ve doğrulama kanıtını değerlendirmez; sonucu siz yazarsınız.
- **Doğrulama bekleme süresi her faaliyet için aynı olmayabilir.** Etkinliğin görülebilmesi için gereken süre
  (yeterli üretim, bir sonraki parti, tekrar denetim) faaliyete göre değişir; 30 gün varsayılanı bir çalışma
  eşiğidir, prosedürünüze göre `--dogrulama` ile ayarlayın.
- **Etkisiz bulunan DÖF otomatik yeniden açılmaz.** Yeni faaliyeti ayrı satır olarak siz eklersiniz; kod iki kaydı
  birbirine bağlamaz.
- **Termin revizyonları görünmez.** Termin ertelenmişse dosyadaki son tarih esas alınır; zamanında aksiyon oranı
  bu yüzden iyimser olabilir.
- **Tek tarih, tek sorumlu.** Birden çok aksiyon adımı olan faaliyetlerde adımları ayrı ayrı izlemez; adım bazlı
  takip için her adımı ayrı satır yazın.
- **Kapsam dışı:** kök neden analizi (5 neden, balık kılçığı), risk değerlendirmesi, müşteriye 8D raporu,
  yönetimin gözden geçirmesi çıktıları.
- **Örnek veri:** Faaliyetler, sorumlular ve tarihler kurgusaldır.

## Testler

- Durum sınırları: termin dün / 15 ve 16 gün sonra, doğrulama beklemesi 30 ve 31 gün, kritik gecikme eşiği.
- Etkin, etkisiz ve sonucu yazılmamış doğrulamalar; kapanma, etkinlik ve zamanında aksiyon oranları, ortalama süreler.
- Tarih sırası tutarsızlıkları, tanınmayan sonuç, sorumlu ve kaynak kırılımı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

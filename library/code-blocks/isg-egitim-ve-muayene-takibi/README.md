# İSG Eğitim ve Muayene Takibi · Kod Bloğu

> İş Sağlığı ve Güvenliği › İş Güvenliği Uzmanı · Workers / Workless

Çalışanların tehlike sınıfına göre **İSG eğitimi** ve **periyodik sağlık muayenesi** tarihlerini takip eder.
Süresi geçenleri ve yaklaşanları tarih sırasıyla bir aksiyon listesinde toplar. Bilgi yenileme ve iş kazası
sonrası ilave eğitimi, işe giriş muayenesini ve ilkyardımcı sayısını da kontrol eder. İnternete bağlanmaz.

## Nasıl hesaplar?

### Eğitim

Dayanak: Çalışanların İş Sağlığı ve Güvenliği Eğitimlerinin Usul ve Esasları Hakkında Yönetmelik (RG 02.04.2026,
sayı 33212). 15.05.2013 tarihli eski yönetmeliğin yerini almıştır.

| Kural | Az tehlikeli | Tehlikeli | Çok tehlikeli | Madde |
|---|---|---|---|---|
| İşe başlama eğitimi (işe başlamadan önce) | en az 2 saat | en az 2 saat | en az 2 saat | 7/5 |
| Temel eğitim (işe başladıktan sonra en geç 3 ay içinde) | en az 8 ders saati | en az 12 | en az 16 | 8/1, 13/1 |
| Tekrar | 3 yılda bir | 2 yılda bir | yılda bir | 14/1 |
| Tekrar eğitimi süresi | en az 8 ders saati | en az 8 | en az 8 | 14/2 |

- **Eğitim döngüsü:** Temel ve tekrar eğitim saatleri tarih sırasıyla birikir.
  - İlk döngü, birikim temel eğitim saatine ulaştığı gün tamamlanır. Sonraki döngüler 8 saate ulaştığı gün
    tamamlanır.
  - Sonraki eğitim tarihi = son tamamlanan döngü + tekrar süresi.
  - Hiç döngüsü olmayan çalışan için son gün işe giriş + 3 aydır.
- **Bilgi yenileme:** "Uzun Ayrılık Dönüşü" doluysa (6 aydan fazla işten uzak kalma), dönüşten önceki 30 gün içinde
  bilgi yenileme eğitimi aranır (md. 18/1).
- **İlave eğitim:** "İş Kazası Dönüşü" doluysa, dönüş tarihinin 30 gün öncesi ile 7 gün sonrası arasında ilave eğitim
  aranır (md. 19/1).
- **İşe başlama eğitimi kontrolü** yalnız yeni yönetmeliğin yürürlüğe girdiği 02.04.2026'dan sonra işe girenlere
  uygulanır.

### Muayene

Dayanak: 6331 sayılı Kanun md. 15 ve İşyeri Hekimi ve Diğer Sağlık Personelinin Görev, Yetki, Sorumluluk ve Eğitimleri
Hakkında Yönetmelik.

- **İşe giriş muayenesi:** işe başlamadan önce yapılmalıdır. Son 1 yıl içinde işe girenlerde kontrol edilir.
- **Periyodik muayene:** en geç 5 / 3 / 1 yılda bir (az tehlikeli / tehlikeli / çok tehlikeli).
  - İşyeri hekiminin belirlediği daha kısa süre "Özel Muayene Periyodu (ay)" sütununa yazılır; ör. gürültü maruziyeti,
    gece çalışması, genç veya gebe çalışan.
  - İki süreden kısa olan kullanılır.
- **Kısıtlı sonuç:** Son muayene sonucu "Çalışabilir" değilse görevlendirme kontrolü için bilgi verilir.
- **İşe dönüş muayenesi:** İş kazası sonrası işe dönüşte muayene kaydı yoksa bilgi verilir. 6331 md. 15'e göre bu
  muayene çalışanın talebi üzerine yapılır.

### İlkyardım

Dayanak: İlkyardım Yönetmeliği.

- İşyeri başına her 20 / 15 / 10 çalışana bir ilkyardımcı gerekir (az / tehlikeli / çok tehlikeli). Oran, işyerindeki
  en yüksek tehlike sınıfına göre alınır.
- İlkyardımcı belgesi son ilkyardım eğitiminden itibaren 3 yıl geçerli sayılır.

### Uyarılar

| Durum | Önem |
|---|---|
| Eğitim veya periyodik muayene gecikmiş; temel eğitim 3 ayda tamamlanmamış | Yüksek |
| Bilgi yenileme / ilave eğitim yok; muayene kaydı hiç yok; işe giriş muayenesi yok | Yüksek |
| İlkyardımcı sayısı yetersiz; ilkyardımcı belgesi bitmiş | Yüksek |
| `--yaklasan` (60) gün içinde eğitim, muayene veya belge sonu | Orta |
| İşe başlama eğitimi eksik veya geç; işe giriş muayenesi işe girişten sonra | Orta |
| Temel eğitim devam ediyor; kısıtlı rapor; işe dönüş muayenesi; ayrılmış çalışana ait kayıt | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek: 26 aktif çalışan, 2 işyeri, rapor tarihi 09.10.2026
python main.py --calisanlar c.xlsx --egitimler e.xlsx --muayeneler m.xlsx --tehlike tehlikeli --bugun 09.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Çalışanlar | Sicil, Ad Soyad, Birim, Görev, İşyeri, İşe Giriş, İşten Çıkış, Tehlike Sınıfı (boşsa `--tehlike`), Özel Muayene Periyodu (ay), Uzun Ayrılık Dönüşü, İş Kazası Dönüşü |
| Eğitimler | Sicil, Tarih, Eğitim Türü (İşe başlama / Temel / Tekrar / Bilgi yenileme / İlave / İlkyardım), Süre (ders saati) |
| Muayeneler | Sicil, Tarih, Muayene Türü (İşe giriş / Periyodik / İşe dönüş), Sonuç |

Ayrılmış çalışanlar (işten çıkış tarihi rapor tarihinden önce) takibe alınmaz.

## Çıktı

`isg_takip.xlsx`:
- `Durum`: çalışan bazında eğitim ve muayene tarihleri ile durumları (renkli).
- `Aksiyon Listesi`: son tarih sırasıyla; boş "Planlanan Tarih" ve "Yapıldı" sütunları.
- `Aylık Plan`: önümüzdeki 6 ayda eğitim ve muayenesi gelen çalışanlar. Gecikmişler bu aya yazılır.
- `Birim Özeti`, `İlkyardımcı`, `Uyarılar`.

## Dikkat

- **Mevzuat:** Süreler yürürlükteki yönetmeliklere göredir. İşyerinin tehlike sınıfı, İş Sağlığı ve Güvenliğine
  İlişkin İşyeri Tehlike Sınıfları Tebliği'ndeki NACE koduna göre belirlenir. Güncel metinleri kontrol edin.
- **Eğitim saatleri** ders saati olarak girilmelidir (45 dakika ders + 15 dakika ara). Eğitimin konu içeriği (Ek-1)
  ve belgelendirme bu pakette kontrol edilmez.
- **Özel eğitimler** (ör. yüksekte çalışma, kapalı alan, forklift operatörlüğü) ve mesleki yeterlilik belgeleri bu
  paketin kapsamında değildir.
- **Sağlık verisi özel nitelikli kişisel veridir** (KVKK md. 6). Muayene sonuçlarını yalnız yetkili kişilerle
  paylaşın; mümkünse "Sonuç" sütununa yalnız çalışabilirlik kararını yazın.
- **Örnek veri:** Örnek çalışanlar ve kayıtlar kurgusaldır.

## Testler

Şunlar test edilir:
- Eğitim döngüsü birikimi (temel 12 / 16 saat, tekrar 8 saat).
- Tehlike sınıfına göre tekrar ve muayene periyotları; özel periyot.
- Yeni girişte temel eğitimin 3 ay kuralı; bilgi yenileme ve ilave eğitim.
- Örnek veride gecikmiş ve yaklaşan kayıtlar, ilkyardımcı yeterliliği, ayrılmış çalışan, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. İSG profesyonelinin
değerlendirmesi yerine geçmez. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

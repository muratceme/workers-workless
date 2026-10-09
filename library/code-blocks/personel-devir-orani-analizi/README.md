# Personel Devir Oranı Analizi · Kod Bloğu

> İnsan Kaynakları › İnsan Kaynakları Müdürü · Workers / Workless

Giriş-çıkış verisinden personel devir (turnover) oranlarını hesaplar. Departman, yönetici ve kıdem bazında
kırılım verir; erken ayrılma sinyallerini çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Ortalama personel:** Dönem başı ve her ay sonu aktif personel sayılarının ortalamasıdır. Ayrılış tarihi son
  çalışma günü kabul edilir; o günün sonunda kişi aktif sayılmaz.
- **Devir oranı:** dönemde ayrılan ÷ ortalama personel. Yıllıklandırılmış oran = dönem oranı × 365 ÷ dönem günü.
  Gönüllü devir yalnız gönüllü ayrılışları sayar.
- **Tutma oranı:** Dönem başındaki personelden dönem sonunda hâlâ çalışanların oranıdır.
- **Ayrılış türü:** "Ayrılış Türü" sütununda Gönüllü / Gönülsüz / Diğer yazıyorsa o kullanılır. Yoksa nedenden
  çıkarılır:

| Neden içinde | Tür |
|---|---|
| istifa, işçi tarafından fesih, evlilik | Gönüllü |
| işveren feshi, işten çıkarma, performans, disiplin, toplu çıkış, işyeri kapanması | Gönülsüz |
| emeklilik, askerlik, ölüm, malulen, belirli süreli sözleşmenin sona ermesi | Diğer |
| tanınmayan | Belirsiz (uyarı) |

- **Kırılımlar:** departman, yönetici, ayrılıştaki kıdem bandı (0–3 ay, 3–12 ay, 1–3 yıl, 3–5 yıl, 5 yıl +) ve aylık
  trend. Kıdem bandında payda, o banttaki ortalama personeldir.
- **Erken ayrılma:**
  - Ayrılanlardan ilk yılını doldurmayanlar.
  - Dönemde işe alınıp 90 günü gözlenebilenlerden 90 gün içinde ayrılanlar, yönetici bazında (`--erken-gun`).

| Sinyal | Önem |
|---|---|
| Aynı yöneticide 2 veya daha fazla erken ayrılış | Yüksek |
| Gönüllü devir şirket ortalamasının 1,5 katını (`--kat`) aşan ve en az 2 gönüllü ayrılışı olan departman / yönetici (ortalama ≥ 3 kişi) | Orta |
| Ayrılanların %30'u veya fazlası ilk yılında | Orta |
| Ortalaması 3 kişinin altındaki grup; türü belirsiz ayrılış | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 63 kişi, 01.01–30.09.2026
python main.py --personel personel.xlsx --donem 01.01.2026 31.12.2026
python main.py --personel personel.xlsx --donem 01.01.2026 30.06.2026 --erken-gun 60 --kat 2
```

Dönem verilmezse son 12 tam ay kullanılır. Dosyada dönemden önce ayrılanlar da bulunabilir; yalnız dönemle kesişen
kayıtlar hesaba girer.

| Sütun | Açıklama |
|---|---|
| Sicil No, İşe Giriş | Zorunlu |
| Ad Soyad, Departman, Pozisyon, Yönetici | Kırılımlar için |
| Ayrılış Tarihi, Ayrılış Nedeni | Çalışmaya devam edenlerde boş |
| Ayrılış Türü | İsteğe bağlı: Gönüllü / Gönülsüz / Diğer |

## Çıktı

`personel_devir_orani.xlsx`:
- `Özet`:
  - dönem başı, sonu ve ortalama personel; işe alınan ve ayrılan
  - devir ve gönüllü devir oranı (yıllıklandırılmış), tutma oranı
  - erken ayrılma
- `Kırılımlar`: departman, yönetici, kıdem.
- `Aylık Trend`: giren / ayrılan grafiğiyle.
- `Erken Ayrılma`: yönetici bazında oran ve kişi listesi.
- `Ayrılanlar`: kıdem bandı, tür (renkli) ve boş "Çıkış Görüşmesi Notu" sütunu.
- `Uyarılar`: boş "Aksiyon" sütunu.

## Dikkat

- **Tanım farkları:** Devir oranının farklı hesaplama yöntemleri vardır; ör. (başı + sonu) ÷ 2 paydası ya da yalnız
  kadrolu personel. Kıyas yaparken aynı tanımı kullanın. Stajyer ve geçici personeli ayrı dosyada analiz etmek
  daha anlamlıdır.
- **Küçük gruplar:** Birkaç kişilik ekipte tek bir ayrılış oranı çok yükseltir. Sinyaller bir inceleme başlangıcıdır;
  yönetici hakkında tek başına değerlendirme gerekçesi değildir.
- **Ayrılış nedeni:** Nedenden tür çıkarımı anahtar kelimeyle yapılır. Kesin sınıflama için "Ayrılış Türü" sütununu
  doldurun. SGK işten çıkış kodlarını kullanıyorsanız kod açıklamasını "Ayrılış Nedeni"ne yazın.
- **Kişisel veri:** Ayrılış nedenleri ve çıkış görüşmesi notları kişisel veridir. Raporu yetkili kişilerle sınırlı
  paylaşın.
- **Örnek veri:** Örnek veri kurgusaldır.

## Testler

Şunlar test edilir:
- Ortalama personel (bağımsız hesapla doğrulandı), devir, yıllıklandırma ve tutma oranı.
- Tür sınıflama; departman, yönetici, kıdem ve aylık kırılımlar.
- Erken ayrılma kohortu ve sinyaller.
- Küçük grup, tarih hatası, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

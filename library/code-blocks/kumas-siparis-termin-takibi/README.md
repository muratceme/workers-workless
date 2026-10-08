# Kumaş Sipariş Termin Takibi · Kod Bloğu

> Tekstil ve Konfeksiyon › Kumaş ve Aksesuar Satın Alma › Kumaş Satın Alma Uzmanı · Workers / Workless

Kumaş siparişlerinin üretim aşamalarını ve termin durumunu takip eder: iplik temini, örme / dokuma, boyahane, apre,
kalite kontrol ve sevk. Gecikmenin bağlı konfeksiyon siparişinin kesim ve sevk tarihine etkisini hesaplar.
İnternete bağlanmaz.

## Nasıl çalışır?

1. **Aşama şablonu:** `asama_sureleri.csv` kumaş tipine göre aşamaları ve standart sürelerini tanımlar. Örme ve
   dokuma için ayrı akış vardır.
2. **Tahmini teslim:** Son tamamlanan aşamanın tarihinden başlar; kalan aşamaların süreleri eklenir (takvim günü).
   - **Gecikmiş aşama:** Süresi geçmiş ama tamamlanmamış aşamanın en erken bugün biteceği varsayılır. "Aşama
     gecikti" uyarısı verilir.
   - **Lab dip onayı:** Boyama, lab dip onayından önce başlamaz. Onay tarihi boşsa boyamanın bugün başlayacağı
     varsayılır ve uyarı verilir.
3. **Termin:** Teslim alınmış siparişte gerçek gecikme, açık siparişte tahmini gecikme hesaplanır. Teslim alınan
   miktar sipariş miktarından %3'ten (`--tolerans`) fazla saparsa eksik veya fazla teslim uyarısı verilir.
4. **Sipariş etkisi:** Konfeksiyon siparişinin kesimi tüm kumaşları (ana kumaş, ribana, garni) gelmeden başlayamaz.
   En geç hazır olan kumaş kritik kumaştır.
   - **Kesim kayması:** Kritik kumaşın hazır tarihine giriş kontrolü (varsayılan 2 gün, `--kontrol-gun`) eklenir.
     Bu tarih planlanan kesimi geçiyorsa kesim kayar.
   - **Sevk:** Kesim ile sevk arasındaki süre sabit kabul edilir ve kayma sevke aynen yansıtılır.

| Uyarı | Önem |
|---|---|
| Kesim / sevk kayıyor | Kritik |
| Termin geçti ve teslim yok; lab dip onayı yokken boyama sırası; eksik teslim | Yüksek |
| Aşama gecikti; tahmini teslim terminden sonra; kesim için 2 gün veya daha az bolluk; kumaş siparişi olmayan konfeksiyon siparişi; aşama tarihleri sırasız | Orta |
| Geç teslim (gerçekleşen); fazla teslim; ara aşama tarihi boş | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek: 6 kumaş siparişi, 4 konfeksiyon siparişi, durum tarihi 08.10.2026
python main.py --kumaslar kumas_siparisleri.xlsx --sureler asama_sureleri.csv --siparisler siparisler.xlsx
python main.py --kumaslar k.xlsx --sureler asama_sureleri.csv --bugun 15.10.2026 --kontrol-gun 3 --tolerans 5
```

| Dosya | Sütunlar |
|---|---|
| Kumaş siparişleri | Kumaş Sipariş No, Tedarikçi, Kumaş, Renk, Tip, Miktar, Birim, Bağlı Sipariş, Sipariş Tarihi, Termin, Lab Dip Onayı, aşama sütunları, Teslim Alınan Miktar, Teslim Tarihi |
| Aşama süreleri | Tip, Aşama, Süre (gün) |
| Konfeksiyon siparişleri (isteğe bağlı) | Sipariş No, Model, Müşteri, Planlanan Kesim Tarihi, Sevk Tarihi |

**Sütun kuralları:**
- **Aşama sütunları:** Başlıkları şablondaki aşama adlarıyla aynı olmalıdır (ör. "Boyama", "Apre / Fikse").
  Hücreye aşamanın tamamlandığı tarih yazılır.
- **Lab Dip Onayı:** Sütun yoksa lab dip kontrolü yapılmaz.

## Çıktı

`kumas_termin_takibi.xlsx`:
- `Sipariş Etkisi`: konfeksiyon siparişi bazında şunları ve boş "Aksiyon / Not" sütununu içerir:
  - kritik kumaş ve kumaş hazır tarihi
  - kesim için en geç tarih
  - kayma ve tahmini sevk
- `Kumaş Durumu`: mevcut aşama, tahmini veya gerçek teslim, termin farkı, durum ve "Tedarikçi Notu" sütunu.
- `Aşama Matrisi`: kumaş × aşama. Gerçekleşen tarihler yeşil, tahmini bitişler gri italik.
- `Tedarikçi Performansı`: zamanında teslim oranı, ortalama gecikme, açık ve gecikmedeki sipariş sayısı.
- `Uyarılar`.

## Dikkat

- **Süreler örnektir:** Aşama süreleri örnektir. Kendi tedarikçi ve boyahane sürelerinizle güncelleyin. Tedarikçinin
  verdiği güncel tarih varsa onu esas alın.
- **Tahminin varsayımları:** Kapasite, boyahane sırası, tekrar boyama (re-dye) ve kalite reddi olasılığı hesaba
  katılmaz.
- **Kesim–sevk süresi:** Kayma sevke aynen yansıtılır. Dikim kapasitesiyle telafi edilebilecek kısım için hat
  planına bakın.
- **Örnek veri:** Örnek tedarikçiler, kumaşlar ve siparişler kurgusaldır.

## Testler

Şunlar test edilir:
- Tahmini teslim (gecikmiş aşama ve lab dip beklemesi).
- Termin ve eksik teslim; kritik kumaş ve kesim / sevk kayması.
- Tedarikçi performansı; veri hataları; Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

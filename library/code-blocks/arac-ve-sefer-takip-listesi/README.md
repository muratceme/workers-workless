# Araç ve Sefer Takip Listesi · Kod Bloğu

> Lojistik ve Taşımacılık › Karayolu Operasyon › Karayolu Operasyon Uzmanı · Workers / Workless

Uluslararası karayolu seferlerini aşama aşama takip eder: yükleme, çıkış gümrüğü, sınır çıkışı, varış gümrüğü ve
teslim. Geciken araçları, teslimi kaçacak seferleri ve araç çakışmalarını listeler. İnternete bağlanmaz.

## Nasıl çalışır?

1. **Aşama süreleri:** `asama_sureleri.csv` varış ülkesine göre her aşamanın bir önceki aşamadan sonra kaç saat
   sürdüğünü tanımlar.
2. **Tahmini teslim:** Son gerçekleşen aşamanın zamanına kalan aşamaların süreleri eklenir.
   - **Gecikmiş aşama:** Beklenen zamanı geçmiş ama gerçekleşmemiş aşamanın (ör. sınır kuyruğu) en erken şimdi
     gerçekleşeceği varsayılır.
   - **Yüklenmemiş sefer:** Planlanan yükleme zamanından başlanır; o zaman geçmişse şimdiden başlanır.
3. **Kontroller:**

| Kontrol | Önem |
|---|---|
| Planlanan teslimi kaçıracak sefer: 24 saatten fazla Yüksek, daha az Orta | Yüksek / Orta |
| Yükleme gecikti: 24 saatten fazla Yüksek | Yüksek / Orta |
| Araç çakışması: aynı çekicinin yeni seferi, önceki seferin tahmini tesliminden önce planlanmış | Yüksek |
| Aşama gecikti (sınır, gümrük); son konum 12 saatten eski (`--konum-saat`) | Orta |
| Aşama zamanları sırasız; varış ülkesi için süre tanımı yok | Orta |
| Geç teslim (gerçekleşen); yüklenmiş seferde CMR numarası yok | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                      # örnek: 8 sefer, durum 08.10.2026 09:00
python main.py --seferler seferler.xlsx --sureler asama_sureleri.csv
python main.py --seferler s.xlsx --sureler asama_sureleri.csv --simdi "10.10.2026 14:00" --konum-saat 6
```

| Dosya | Sütunlar |
|---|---|
| Seferler | Sefer No, Çekici Plaka, Dorse Plaka, Şoför, Müşteri, Varış Ülke, Varış Şehri, Sınır Kapısı, Planlanan Yükleme, Planlanan Teslim, Yükleme, Çıkış Gümrüğü, Sınır Çıkış, Varış Gümrüğü, Teslim, CMR No, Son Konum, Son Konum Zamanı |
| Aşama süreleri | Varış Ülke, Aşama, Süre (saat) |

**Doldurma kuralları:**
- **Zaman biçimi:** "GG.AA.YYYY SS:DD".
- **Aşama sütunları:** Gerçekleştiği zaman yazılır, gerçekleşmediyse boş bırakılır.
- **Yurt içi seferler:** Gümrük ve sınır aşamaları süre tablosunda tanımlanmazsa atlanır.
- **Tanımsız ülke:** Süre tablosunda "Varış Ülke" olarak `*` yazılan satırlar, tanımsız ülkeler için varsayılan
  süredir.

## Çıktı

`sefer_takip.xlsx`:
- `Seferler`: gecikecekler önce. Mevcut aşama, tahmini veya gerçek teslim, fark (saat), son konum ve boş
  "Operasyon Notu" sütunu.
- `Araç Durumu`: çekici bazında aktif sefer, son konum, müsait olacağı zaman (tahmini teslim) ve sıradaki seferler.
- `Aşama Matrisi`: sefer × aşama. Gerçekleşen zamanlar yeşil, tahminler gri italik.
- `Performans`: müşteri ve varış ülkesi bazında zamanında teslim oranı.
- `Uyarılar`.

## Dikkat

- **Süreler örnektir:** Aşama süreleri örnektir. Kendi güzergâh geçmişinizle güncelleyin.
- **Tahminin kapsamı:** Sınır kapılarındaki bekleme mevsime ve güne göre çok değişir. Tahmin şunları hesaba katmaz:
  - şoför sürüş ve dinlenme süreleri (AETR)
  - hafta sonu sürüş yasakları
  - feribot ve tren bağlantıları

  Güncel bilgi için şoför ve gümrük müşaviriyle teyit edin.
- **Araç çakışması:** Kontrol yalnızca teslim zamanına bakar. Boşaltma sonrası dönüş süresini planlayıcı ayrıca
  değerlendirmelidir.
- **Örnek veri:** Örnek seferler, plakalar ve müşteriler kurgusaldır.

## Testler

Şunlar test edilir:
- Tahmini teslim; gecikmiş aşama ve gecikmiş yükleme.
- Teslim gecikmesi, araç çakışması, eski konum, CMR ve veri hataları.
- Araç durumu, performans, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

# Tedarikçi Performans Değerlendirme · Kod Bloğu

> Satın Alma › Satın Alma Uzmanı · Workers / Workless

Teslimat kayıtlarından **tedarikçi karnesi** hazırlar. Zamanında teslim, kalite reddi, miktar uyumu ve fiyat uyumunu
puanlar; tedarikçileri A/B/C/D olarak sınıflar, çeyreklik eğilimi gösterir ve tedarikçiye gönderilebilecek karne metnini
yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Sipariş satırı:** Sipariş No + Malzeme. Kısmi teslimatlar ayrı satırlarda girilir ve toplanır.
- **Tamamlanma tarihi:** Toplam teslim, sipariş miktarının (1 − miktar toleransı) kadarına ulaştığı teslimatın tarihi.
  - Eksik teslim edilmiş satır "Satır Durumu = Açık" değilse son teslimatla kapanmış sayılır; miktar uyumu "Hayır" olur.
  - Termini geçmiş, tamamlanmamış açık satır **geciken** sayılır. Gecikme rapor tarihine kadar hesaplanır.
  - Termini gelmemiş açık satır değerlendirmeye girmez.

| Kriter | Puan (0–100) | Varsayılan ağırlık |
|---|---|---|
| Zamanında teslim | Tamamlanma ≤ termin + `--tolerans-gun` (0) olan satırların oranı | 40 |
| Kalite | Retsiz teslimatların (lot) oranı; ayrıca ret oranı PPM olarak gösterilir | 35 |
| Fiyat uyumu | Fatura birim fiyatı ≤ sipariş fiyatı × (1 + %0,5) olan satırların oranı | 15 |
| Miktar uyumu | Teslim toplamı sipariş miktarının ±%5'i içinde olan satırların oranı | 10 |

- **Toplam puan:** Kriter puanlarının ağırlıklı ortalaması. Verisi olmayan kriter (ör. fatura fiyatı yok) ağırlıktan
  çıkarılır ve uyarıda belirtilir.
- **Sınıf:** A ≥ 90, B ≥ 75, C ≥ 60, altı D (`--siniflar`). Değerlendirilen satır `--min-satir` (3) altındaysa sınıf
  verilmez.
- **Fazla faturalama:** (fatura fiyatı − sipariş fiyatı) × teslim miktarı. Tolerans içindeki farklar da tutar olarak
  gösterilir.
- **Çeyreklik eğilim:** Termin tarihine göre çeyrek puanları. Son çeyrek bir öncekine göre `--dusus` (10) puan veya daha
  fazla düştüyse uyarı verilir.

| Kontrol | Önem |
|---|---|
| D sınıfı tedarikçi | Yüksek |
| Termini geçmiş açık sipariş | Yüksek |
| Retsiz lot oranı %90'ın altında (8D / DÖF isteyin) | Orta |
| Sipariş fiyatının üzerinde faturalama | Orta |
| Önceki çeyreğe göre puan düşüşü | Orta |
| Yetersiz veri; verisi olmayan kriter | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 6 tedarikçi, Ocak–Eylül 2026, rapor tarihi 09.10.2026
python main.py --teslimatlar teslimatlar.xlsx --bugun 09.10.2026
python main.py --teslimatlar t.xlsx --agirlik teslim=50 kalite=30 fiyat=10 miktar=10 --tolerans-gun 2 --siniflar 85,70,50
```

| Sütun | Açıklama |
|---|---|
| Sipariş No, Tedarikçi, Malzeme | Satır anahtarı |
| Sipariş Tarihi, Termin Tarihi | Termin zorunludur |
| Teslim Tarihi, Teslim Miktarı | Her kısmi teslimat ayrı satır; teslim yoksa boş |
| Ret Miktarı (veya Kabul Miktarı) | Giriş kalite kontrolünde reddedilen |
| Sipariş Birim Fiyatı, Fatura Birim Fiyatı | Fiyat uyumu için |
| Satır Durumu | İsteğe bağlı: Açık / Kapalı |

## Çıktı

`tedarikci_karnesi.xlsx`:
- `Karne`: puana göre sıralı; boş "Aksiyon" sütunu.
- `Çeyreklik Trend`: grafikli.
- `Geciken Siparişler`, `Kalite Retleri` (boş DÖF / 8D no sütunu), `Fiyat Farkları`.
- `Sipariş Satırları`: her satırın tamamlanma tarihi, durumu ve kriter sonuçları.
- `Karne Metinleri`: tedarikçiye gönderilebilecek kısa değerlendirme.
- `Uyarılar`.

## Dikkat

- **Ağırlıklar ve sınıf sınırları** örnektir. Şirketin tedarikçi değerlendirme prosedürüne (ör. ISO 9001 kapsamındaki
  dış kaynak değerlendirmesi) göre ayarlayın.
- **Termin değişiklikleri:** Termin tedarikçiyle yazılı olarak değiştirildiyse güncel termini girin. Aksi hâlde gecikme
  haksız yere yüksek görünür.
- **Ret nedeni:** Tedarikçiden kaynaklanmayan ret (ör. yanlış sipariş) varsa ret miktarına yazmayın.
- **Karne metni** göndermeden önce gözden geçirilmelidir.
- **Örnek veri:** Örnek tedarikçiler, siparişler ve fiyatlar kurgusaldır.

## Testler

Şunlar test edilir:
- Bir tedarikçinin tüm kriterleri elle hesaplanmış değerlerle.
- Kısmi teslimle tamamlanma, eksik kapama, açık gecikmiş ve termini gelmemiş satır.
- Tolerans sınırları ve verisi olmayan kriterde ağırlığın yeniden dağıtılması.
- Sınıflar, uyarılar, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

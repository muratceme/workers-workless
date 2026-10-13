# İş İlerleme Raporu · Kod Bloğu

> Şantiye › Şantiye Şefi · Workers / Workless

İş programındaki tarihler ve sahadan gelen gerçekleşen yüzdelerden **planlanan-gerçekleşen ilerlemeyi, S-eğrisini ve
geciken iş kalemlerini** raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kalemin planlanan yüzdesi | (rapor tarihi − başlangıç + 1) ÷ (bitiş − başlangıç + 1) × 100, 0–100 arasında |
| Kalemin payı | Ağırlık ÷ ağırlıklar toplamı × 100 |
| Proje ilerlemesi | Σ pay × kalem yüzdesi ÷ 100 (planlanan ve gerçekleşen ayrı) |
| Sapma (puan) | Gerçekleşen % − planlanan % |
| Projeye etkisi (puan) | Pay × sapma ÷ 100 |
| Gecikme (gün, kaba) | (Planlanan % − gerçekleşen %) ÷ 100 × süre |

Durumlar: **Tamamlandı**, **Süresi geçti** (bitiş tarihi geçmiş, iş bitmemiş), **Başlamadı (gecikmede)**,
**Gecikmede** (sapma −tolerans puanından kötü), **Önde**, **Planında**, **Başlamadı** (başlangıcı gelmemiş).

S-eğrisi, ay (veya hafta) sonlarındaki planlanan kümülatif proje ilerlemesidir; rapor tarihinde gerçekleşen değerle
birlikte verilir.

**Örnek:** Temel betonarme 21.09–20.10 (30 gün), ağırlık %20, gerçekleşen %60; rapor tarihi 12.10. Geçen gün 22 →
planlanan %73,33; sapma −13,33 puan; projeye etkisi −2,67 puan; gecikme ≈ 4 gün.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek veri, rapor tarihi 12.10.2026
python main.py --kalemler is_kalemleri.xlsx --tarih 12.10.2026
python main.py --kalemler is_kalemleri.xlsx --tolerans 3 --donem hafta
```

Sütunlar: İş Kalemi, Başlangıç, Bitiş, Ağırlık (yüzde, pursantaj veya bütçe tutarı), Gerçekleşen % (boşsa 0).
`--tarih` verilmezse bugünün tarihi kullanılır.

## Çıktı

`cikti/is_ilerleme_raporu.xlsx`:

- `Özet`: planlanan ve gerçekleşen proje ilerlemesi, sapma, durum.
- `İş Kalemleri`: kalem bazında planlanan, gerçekleşen, sapma, projeye etkisi, durum.
- `S-Eğrisi`: dönem sonu planlanan kümülatif yüzdeler ve grafik.
- `Gecikenler`: proje ilerlemesini en çok geri çeken kalemler.
- `Uyarılar`: süresi geçen, başlamayan ve geciken kalemler; okunamayan satırlar.

## Dikkat

- **Doğrusal dağılım bir varsayımdır.** Kod, işin süreye eşit yayıldığını kabul eder. Başı veya sonu yoğun kalemlerde
  (örneğin kalıp-demir-beton döngüsü) planlanan yüzde gerçek programınızdan farklı çıkar.
- **Takvim günü kullanılır.** Hafta sonu, resmi tatil ve hava koşulu nedeniyle çalışılmayan günler ayrılmaz.
- **"Gecikenler" kritik yol değildir.** Bolluğu olan bir kalemin gecikmesi proje bitişini etkilemeyebilir; bitiş
  tarihine etki için "İş Programı Kritik Yol Analizi" paketini kullanın.
- **Gerçekleşen yüzde ölçüme dayanmalı.** Yüzdeyi metraj veya hakedişten alın; göz kararı yüzdeler raporu anlamsız kılar.
- **Ağırlık seçimi sonucu değiştirir.** Bütçe ağırlığı ile adam-saat ağırlığı farklı proje yüzdesi verir; sözleşmenizdeki
  pursantajı kullanın.
- **Örnek veri:** İş kalemleri, tarihler ve yüzdeler kurgusaldır.

## Testler

- Kalem bazında planlanan yüzde, sapma, etki, gecikme ve durumlar; proje toplamları.
- Aylık ve haftalık S-eğrisi noktaları; tolerans sınırları; başlamamış, süresi geçmiş ve önde kalemler.
- Okunamayan satır, komut satırı hataları; örnek veri uçtan uca (grafik dahil).

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

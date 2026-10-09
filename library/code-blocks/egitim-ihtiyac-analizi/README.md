# Eğitim İhtiyaç Analizi · Kod Bloğu

> İnsan Kaynakları › Eğitim ve Gelişim Uzmanı · Workers / Workless

Pozisyon yetkinlik matrisini, yönetici ve öz değerlendirmelerini, performans sonuçlarını ve eğitim taleplerini
birleştirir. Departman bazında önceliklendirilmiş eğitim ihtiyacı ve bir eğitim planı taslağı çıkarır. İnternete
bağlanmaz.

## Nasıl hesaplar?

- **Mevcut seviye:** Yönetici puanıdır. Yönetici puanı yoksa öz değerlendirme kullanılır ve bu not edilir.
- **Açık:** gerekli seviye − mevcut seviye (0'ın altına inmez).
- **İhtiyaç puanı:** açık × kritiklik ağırlığı + düşük performans + talep.

| Bileşen | Değer |
|---|---|
| Kritiklik ağırlığı | Yüksek 3, Orta 2, Düşük 1 |
| Düşük performans (C / D veya 1–2) | +1 |
| Çalışan bu eğitimi talep etmiş | +1 |
| **Öncelik** | puan ≥ 6 → 1, 3–5 → 2, 1–2 → 3 |

- **Öncelik listesi:** Departman × yetkinlik bazında açığı olan kişi sayısı, ortalama açık ve toplam puan
  hesaplanır; liste puana göre sıralanır.
- **Eğitim planı:** Aynı önerilen eğitime ihtiyacı olanlar gruplanır. Katılımcı sayısı `--min-grup` (varsayılan 5)
  ve üstündeyse grup eğitimi, altındaysa bireysel yöntem önerilir: e-öğrenme, mentorluk veya iş başında eğitim.
- **Algı farkı:** Öz değerlendirme yönetici puanından 2 veya daha fazla yüksekse "kör nokta", düşükse "fark
  edilmemiş güç" olarak işaretlenir. Bunlar gelişim görüşmesinde konuşulacak konulardır.
- **Talepler:** Talep, önerilen eğitim veya yetkinlik adıyla eşleştirilir. Sonuç şu üç durumdan biridir: talep ve
  ihtiyaç örtüşüyor; talep var ama ölçülen açık yok; matristeki eğitimlerle eşleşmedi.
- **Uyarılar:**
  - Matriste tanımı olmayan pozisyon veya yetkinlik.
  - Değerlendirilmemiş yetkinlik.
  - Performans kaydı olmayan çalışan.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 12 çalışan, 3 pozisyon, 44 değerlendirme
python main.py --matris matris.xlsx --degerlendirmeler d.xlsx --performans p.xlsx --talepler t.xlsx
python main.py --matris m.csv --degerlendirmeler d.csv --min-grup 8
```

| Dosya | Sütunlar |
|---|---|
| Yetkinlik matrisi | Pozisyon, Yetkinlik, Gerekli Seviye (1–5), Kritiklik (Yüksek / Orta / Düşük), Önerilen Eğitim |
| Değerlendirmeler | Sicil No, Ad Soyad, Departman, Pozisyon, Yetkinlik, Yönetici Puanı, Öz Değerlendirme (1–5) |
| Performans (isteğe bağlı) | Sicil No, Performans (A–D veya 1–5) |
| Talepler (isteğe bağlı) | Sicil No, Talep Edilen Eğitim |

Değerlendirmeler bir anket aracından (ör. form çıktısı) her satırda bir kişi × yetkinlik olacak biçimde alınabilir.

## Çıktı

`egitim_ihtiyac_analizi.xlsx`:
- `Öncelik Listesi`: departman × yetkinlik; boş "Karar" sütunu.
- `Eğitim Planı`: katılımcı sayısı, önerilen yöntem, katılımcılar; boş "Planlanan Tarih" ve "Bütçe" sütunları.
- `Bireysel İhtiyaçlar`: kişi × yetkinlik; açık, puan, öncelik (renkli); boş "Onay" sütunu.
- `Departman Özeti`: yetkinlik karşılama oranı ve öncelik dağılımı.
- `Algı Farkları`, `Talepler`, `Uyarılar`.

## Dikkat

- **Matris sizindir:** Yetkinlikler, gerekli seviyeler ve kritiklik pozisyon tanımlarınıza göre belirlenmelidir.
  Örnek matris kurgusaldır.
- **Ağırlıklar bir yöntem önerisidir:** Puan ve öncelik eşikleri karar destek içindir. Nihai planı bütçe, iş
  takvimi ve yönetici görüşüyle birlikte yapın.
- **Yasal zorunlu eğitimler ayrıdır:** İSG eğitimleri gibi yasal zorunlu eğitimler ihtiyaç analizinden bağımsız
  olarak mevzuattaki sürelerde verilmelidir. Bu paket onları takip etmez.
- **Kişisel veri:** Performans ve değerlendirme sonuçları kişisel veridir. Raporu yetkili kişilerle sınırlı
  paylaşın.

## Testler

Şunlar test edilir:
- Puan ve öncelik: kritiklik, performans, talep; yalnız öz değerlendirme; geçersiz puan.
- Departman önceliği ve eğitim planı (grup / bireysel, `--min-grup`).
- Algı farkları, talep eşleştirme ve uyarılar.
- Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

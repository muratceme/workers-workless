# Eğitim Değerlendirme Raporu · Kod Bloğu

> İnsan Kaynakları › Eğitim ve Gelişim Uzmanı · Workers / Workless

Katılım, ön test, son test ve memnuniyet anketi verisinden **eğitim bazında etkinlik raporu** üretir; eşiklerin
altında kalan eğitimleri ve takip edilmesi gereken kişileri listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Katılım % | Katılan ÷ davet edilen × 100 |
| Ön / son test ortalaması | Puanı girilmiş katılımcıların ortalaması |
| Gelişim (puan) | Hem ön hem son testi olanların (son − ön) ortalaması |
| Başarı % | Son testi `--gecme` (70) ve üzeri olan ÷ son testi olan × 100 |
| Memnuniyet | Anket puanlarının ortalaması (`--olcek` 5 üzerinden) |

İşaretler: katılım `--katilim-esik` (%80) altında · başarı `--basari-esik` (%70) altında · memnuniyet
`--memnuniyet-esik` (3,5) altında · gelişim sıfır veya negatif · katılım var ama hiç ölçüm yok.

"Katıldı" sütunu boşsa, test veya anket puanı girilmiş kişi katılmış sayılır.

**Örnek:** 5 davetli, 4 katılan → katılım %80. Ön testler 40, 60, 50; son testler 80, 70, 65, 90. Gelişim, iki testi
de olan üç kişiden: (40 + 10 + 15) ÷ 3 = 21,67 puan. Son testte 70 ve üzeri 3 kişi → başarı %75.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --katilim egitim_katilim.xlsx
python main.py --katilim egitim_katilim.xlsx --gecme 60 --olcek 10 --memnuniyet-esik 7 --katilim-esik 90
```

Zorunlu sütunlar: Eğitim, Katılımcı. Diğerleri: Departman, Katıldı (Evet / Hayır), Ön Test, Son Test (0–100),
Memnuniyet (0 – ölçek).

## Çıktı

`cikti/egitim_degerlendirme_raporu.xlsx`:

- `Eğitim Özeti`: eğitim bazında tüm göstergeler ve işaretler; sorunlu eğitimler üstte.
- `Departman Özeti`: departman bazında katılım, test ortalamaları, başarı ve memnuniyet.
- `Takip Listesi`: katılmayanlar ve geçme puanının altında kalanlar.
- `Uyarılar`: eşiklerin altındaki eğitimler, düşük anket yanıt oranı, çelişkili ve okunamayan satırlar.

## Dikkat

- **Rapor tepkiyi ve öğrenmeyi ölçer, etkiyi değil.** Memnuniyet ve test puanı, eğitimin iş başındaki davranışı
  veya iş sonuçlarını değiştirdiğini göstermez; bunun için sonradan gözlem ve performans verisi gerekir.
- **Ön ve son test aynı zorlukta olmalı.** Farklı sorularla yapılan testlerde "gelişim" yanıltır.
- **Az katılımcıda yüzdeler oynaktır.** İki kişilik eğitimde bir kişi başarıyı %50 değiştirir.
- **Memnuniyet anketi isimsiz toplanıyorsa** puanı kişiyle eşleştirmeyin; anket puanlarını ayrı satırlarda,
  katılımcı adını genel bir etiketle girin.
- **Eşikler varsayımdır.** Kurumunuzun hedeflerine göre değiştirin.
- **Örnek veri:** Eğitimler, kişiler ve puanlar kurgusaldır.

## Testler

- Katılım, ortalamalar, gelişim, başarı ve memnuniyet; genel ve departman özetleri.
- Her işaretin eşik sınırı; hiç ölçümü olmayan eğitim, düşük anket yanıt oranı, takip listesi.
- "Katıldı" boşken çıkarım, çelişkili kayıt, ölçek dışı puan, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

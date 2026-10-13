# İşe Alım Süreç Raporu · Kod Bloğu

> İnsan Kaynakları › İşe Alım Uzmanı · Workers / Workless

Aday takip listesinden açık pozisyonların **aşama dağılımını, bekleme sürelerini, kaynak kanal kırılımını ve işe alım
süresini** (time-to-hire) raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| İşe alım süresi | "İşe Alındı" aşama tarihi − başvuru tarihi (gün); pozisyon ve genel ortalama |
| Pozisyon açık gün | (Kapanış tarihi, yoksa bugün) − açılış tarihi |
| Doldurma süresi | Kapanmış pozisyonların açık gün ortalaması |
| Aşamada bekleme | Bugün − adayın son aşama tarihi; `--bekleme` (7) günü aşanlar listelenir |
| Mülakata ulaşma % | Mülakat, Teklif veya İşe Alındı aşamasındaki aday ÷ kaynaktan gelen aday × 100 |
| Kaynak dönüşümü % | Kaynaktan işe alınan ÷ kaynaktan gelen aday × 100 |

Aşamalar: **Başvuru, Ön Görüşme, Mülakat, Teklif** (aktif), **İşe Alındı, Olumsuz, Çekildi**. "Ön eleme", "Red",
"Kabul", "Vazgeçti" gibi yaygın yazımlar bu aşamalara eşlenir.

**Örnek:** Aday 01.09'da başvurmuş, 01.10'da işe alınmış → işe alım süresi 30 gün. Pozisyon 25.08'de açılmış, rapor
tarihi 13.10 → 49 gündür açık. Mülakat aşamasındaki adayın son işlemi 01.10 → 12 gündür bekliyor.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, rapor tarihi 13.10.2026
python main.py --adaylar adaylar.xlsx
python main.py --adaylar adaylar.xlsx --pozisyonlar pozisyonlar.xlsx --tarih 13.10.2026 --bekleme 5
```

- Adaylar: Pozisyon, Aday, Başvuru Tarihi, Aşama (zorunlu); Kaynak, Aşama Tarihi (boşsa başvuru tarihi alınır).
- Pozisyonlar (isteğe bağlı): Pozisyon, Açılış Tarihi (zorunlu); Departman, Kapanış Tarihi, Sorumlu.

## Çıktı

`cikti/ise_alim_surec_raporu.xlsx`:

- `Pozisyon Özeti`: pozisyon bazında aşama dağılımı, aktif ve bekleyen aday, açık gün, ortalama işe alım süresi.
- `Kaynak Analizi`: kanal bazında aday, mülakata ulaşma ve dönüşüm oranı, ortalama işe alım süresi.
- `Bekleyen Adaylar`: aktif aşamada uzun süredir işlem görmeyen adaylar.
- `Uyarılar`: adayı veya aktif adayı olmayan açık pozisyonlar, bekleyen adaylar, okunamayan satırlar.

## Dikkat

- **Rapor yalnız şimdiki aşamayı bilir.** Olumsuz sonuçlanan adayın hangi aşamada elendiği dosyada yoksa huni
  (aşamadan aşamaya geçiş) analizi yapılamaz; "mülakata ulaşma" oranı bu yüzden olduğundan düşük çıkar.
- **İşe alım süresi tanımı kurumdan kuruma değişir.** Kod başvurudan işe alındı kaydına kadar olan süreyi ölçer;
  teklif kabulünü veya işe başlama gününü esas alıyorsanız aşama tarihini ona göre girin.
- **Az sayıda işe alımda ortalama yanıltır.** Tek bir uzun süreç ortalamayı kaydırır.
- **Aday verisi kişisel veridir.** Raporu yalnız süreçte görevli kişilerle paylaşın; adaylara yaptığınız aydınlatmaya
  ve saklama sürelerinize uyun.
- **Kaynak kalitesi yalnız dönüşümle ölçülmez.** İşe alınanın performansı ve kalıcılığı bu raporda yoktur.
- **Örnek veri:** Adaylar, pozisyonlar ve kaynaklar kurgusaldır.

## Testler

- Pozisyon bazında aşama sayıları, açık gün, işe alım ve doldurma süresi; genel toplamlar.
- Kaynak kırılımı, bekleyen adaylar ve `--bekleme` eşiği; pozisyon dosyası olmadan çalışma.
- Adayı olmayan açık pozisyon, ileri tarihli aşama, tanınmayan aşama, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

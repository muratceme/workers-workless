# MTBF ve MTTR Analizi · Kod Bloğu

> Bakım Onarım › Bakım Mühendisi · Workers / Workless

Arıza kayıtlarından makine bazında **arızalar arası ortalama süreyi (MTBF), ortalama onarım süresini (MTTR) ve
kullanılabilirliği** hesaplar; arıza tiplerini süre kaybına göre sıralar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Arıza süresi (saat) = bitiş − başlangıç            (yoksa "Süre (dk)" ÷ 60)
MTTR                = toplam arıza süresi ÷ arıza sayısı
Çalışma süresi      = planlı çalışma süresi − toplam arıza süresi
MTBF                = çalışma süresi ÷ arıza sayısı
Kullanılabilirlik % = çalışma süresi ÷ planlı çalışma süresi × 100     ( = MTBF ÷ (MTBF + MTTR) )
```

- **Örnek:** planlı çalışma 480 saat, 4 arıza, toplam 8 saat → MTTR 2,00 saat; çalışma 472 saat → MTBF 118,00 saat;
  kullanılabilirlik %98,33.
- **Planlı çalışma süresi** makine dosyasından okunur; dosyada olmayan makineler için `--planli-saat` kullanılır.
  İkisi de yoksa yalnız MTTR hesaplanır.
- **Arızası olmayan makine** (makine dosyasında var, arıza kaydı yok): MTBF "arıza yok", kullanılabilirlik %100.
- **Devam eden arıza** (bitişi boş) hesaba alınmaz ve uyarı olarak listelenir.
- **Genel MTBF** yalnız planlı süresi bilinen makinelerin toplamından hesaplanır.
- `--mtbf-hedef` ve `--mttr-hedef` verilirse hedefi tutmayan makineler işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --arizalar ariza_kayitlari.xlsx --makineler makineler.xlsx --mtbf-hedef 150 --mttr-hedef 4
python main.py --arizalar ariza_kayitlari.xlsx --planli-saat 480
```

| Dosya | Sütunlar |
|---|---|
| Arıza kayıtları | Makine (zorunlu); Başlangıç ve Bitiş (`GG.AA.YYYY SS:DD`) veya Süre (dk); Arıza Tipi |
| Makineler (isteğe bağlı) | Makine, Planlı Çalışma Süresi (saat) — arıza dosyasının kapsadığı dönem için |

## Çıktı

`cikti/mtbf_mttr_analizi.xlsx`:

- `Makine Bazında`: arıza sayısı, arıza süresi, MTBF, MTTR, kullanılabilirlik, en çok süre kaybettiren arıza tipi; boş "Aksiyon" sütunu.
- `Arıza Tipi Bazında`: arıza sayısı, toplam süre, MTTR, süre payı ve birikimli pay.
- `Uyarılar`: kapanmamış arızalar, planlı süresi olmayan makineler, okunamayan kayıtlar.

## Dikkat

- **Tanım farkı sonucu değiştirir.** Burada MTBF, çalışma süresinin arıza sayısına bölümüdür. Kuruluşunuz takvim
  süresini veya iki arıza başlangıcı arasındaki süreyi kullanıyorsa sonuçlar farklı çıkar; karşılaştırmadan önce
  tanımı eşleştirin.
- **Arıza süresi kayıttaki duruştur.** Parça veya teknisyen bekleme süresi kayda dahilse MTTR yalnız onarım süresini
  değil toplam duruşu gösterir.
- **Planlı çalışma süresi doğru olmalı.** Planlı bakım, mola ve üretim planı olmayan zaman bu süreye dahil edilmemelidir.
- **Az arızayla istatistik zayıftır.** Bir iki arızalı makinede MTBF tek bir olayla büyük oynar; uzun dönem verin.
- **Aynı arızanın tekrar kayıtları** (kısa aralıkla kapatılıp yeniden açılan) arıza sayısını şişirir, MTBF'yi düşürür.
- **Örnek veri:** Makineler ve arıza kayıtları kurgusaldır.

## Testler

- MTBF, MTTR ve kullanılabilirlik; MTBF ÷ (MTBF + MTTR) özdeşliği.
- Arızasız makine, planlı süresi olmayan makine, arıza süresi planlı süreyi aşan makine.
- Hedef işaretleri, arıza tipi sıralaması, genel göstergeler; tarih-saat okuma.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

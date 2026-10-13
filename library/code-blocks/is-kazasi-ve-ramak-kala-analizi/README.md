# İş Kazası ve Ramak Kala Analizi · Kod Bloğu

> İş Sağlığı ve Güvenliği › İş Güvenliği Uzmanı · Workers / Workless

Kaza ve ramak kala kayıtlarınızdan **kaza sıklık ve ağırlık oranlarını** hesaplar; olayları bölüm, neden, ay,
haftanın günü, saat dilimi ve vardiyaya göre döker. Yıllık değerlendirme raporu ve kurul toplantısı için tablo
hazırlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kaza sıklık oranı | İş kazası sayısı × 1.000.000 ÷ toplam çalışma saati |
| Kaza ağırlık oranı | Kayıp iş günü × 1.000 ÷ toplam çalışma saati |
| 1.000 çalışan başına kaza | İş kazası sayısı × 1.000 ÷ ortalama çalışan sayısı |
| Ramak kala / kaza | Ramak kala sayısı ÷ iş kazası sayısı |
| Kaza başına kayıp gün | Kayıp iş günü ÷ iş kazası sayısı |
| Neden Pareto | Nedenler olay sayısına göre azalan sırada; pay % ve kümülatif % |

Oranlar yalnız "İş Kazası" kayıtlarından hesaplanır. Bölüm dosyası verirseniz aynı oranlar bölüm bazında da çıkar.

**Örnek:** 7 iş kazası, 64 kayıp gün, 480.000 çalışma saati → sıklık oranı 7 × 1.000.000 ÷ 480.000 = **14,58**;
ağırlık oranı 64 × 1.000 ÷ 480.000 = **0,13**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --olaylar olaylar.xlsx --bolumler bolumler.xlsx
python main.py --olaylar olaylar.xlsx --saat 480000 --calisan 220
python main.py --olaylar olaylar.xlsx --saat 480000 --agirlik-carpan 1000000 --tekrar 2
```

Zorunlu sütunlar: Tarih ve Tür. Tür "İş Kazası" ya da "Ramak Kala" olmalıdır. Çalışma saati verilmezse sayımlar ve
kırılımlar üretilir, oranlar boş kalır.

## Çıktı

`cikti/is_kazasi_ramak_kala_analizi.xlsx`:

- `Özet`: olay sayıları, kayıp gün ve oranlar.
- `Bölümler`: bölüm bazında kaza, ramak kala, kayıp gün ve (bölüm dosyası varsa) oranlar.
- `Neden Pareto`: olayların çoğunu oluşturan nedenler; kümülatif %80'e kadar olanlar renklidir.
- `Aylar`, `Haftanın Günü`, `Saat Dilimi`, `Vardiya`: zaman kırılımları.
- `Olaylar` ve `Uyarılar` (tekrarlayan neden, bölüm bilgisi eksikliği, hiç ramak kala kaydı olmaması).

## Dikkat

- **Çarpanlar kaynağa göre değişir.** Sıklık oranı için 1.000.000, ağırlık oranı için 1.000 çarpanı varsayılandır;
  bazı kaynaklar ve resmî istatistikler farklı çarpan ya da farklı tanım kullanır. Bir sektör ortalamasıyla
  kıyaslayacaksanız o kaynağın çarpanını `--siklik-carpan` ve `--agirlik-carpan` ile verin; yoksa karşılaştırma
  yanlış olur.
- **Çalışma saati sizin girdinizdir.** Fiilen çalışılan saat (fazla mesai dahil, izin ve rapor hariç) olmalıdır;
  çalışan sayısı × yıllık saat gibi kaba tahmin oranı bozar. Alt işveren çalışanlarını olaylara dahil ediyorsanız
  saatlerini de dahil edin.
- **Kayıp gün tanımını tutarlı kullanın.** Kaza gününün sayılıp sayılmayacağı, ölüm ve sürekli iş göremezlik için
  hangi gün karşılığının yazılacağı kullandığınız yönteme bağlıdır; kod yazdığınız sayıyı toplar.
- **Küçük sayılar yanıltır.** 20 kişilik bir bölümde tek kaza oranı sıçratır; oranları olay sayılarıyla birlikte
  okuyun. Ramak kala sayısının düşük olması güvenli demek değildir, bildirilmiyor olabilir.
- **Yasal bildirim bu paketin dışındadır.** İş kazasının kuruma bildirilmesi, kayıt ve tutanak yükümlülükleri
  ayrıca yerine getirilir. Kök neden analizi de (5 neden, balık kılçığı) insan işidir; kod yalnız sayar.
- **Kişisel veri:** Olay açıklamalarına kazazedenin adını ve sağlık bilgisini yazmayın; rapor paylaşılabilir.
- **Örnek veri:** Olaylar, bölümler ve saatler kurgusaldır.

## Testler

- Oranlar ve çarpan değişimi, kayıp günü boş ve sıfır kaza, ramak kala olaylarının orana girmemesi.
- Neden Pareto sırası ve kümülatif yüzde, tekrarlayan neden eşiği, bölüm oranları ve olaysız bölüm.
- Ay, haftanın günü, saat dilimi (gece yarısı sınırı) ve vardiya kırılımı, saat yazımları, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

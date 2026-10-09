# Dönem Sonu Kur Değerleme · Kod Bloğu

> Muhasebe › Genel Muhasebe Uzmanı · Workers / Workless

Döviz cinsinden hesapların bakiyelerini dönem sonu kurlarıyla değerler. Kur farkını hesaplar ve yevmiye kaydı
önerisi hazırlar. İnternete bağlanmaz; kurları siz girersiniz.

## Nasıl hesaplar?

- **Değerlenmiş TL:** döviz bakiye × değerleme kuru. Hesap satırında "Kur" verilirse o kullanılır.
- **Kur farkı:** değerlenmiş TL − kayıtlı TL. Borç bakiyesi pozitif, alacak bakiyesi negatif alınır.

| Fark | Kayıt |
|---|---|
| > 0 | İlgili hesap borç / 646 Kambiyo Kârları alacak |
| < 0 | 656 Kambiyo Zararları borç / ilgili hesap alacak |

Örneğin alacak bakiyeli bir döviz kredisinde (300) kur artışı borcu büyütür ve zarar yazar. Kâr ve zarar hesap
kodları `--kar-hesabi` ve `--zarar-hesabi` ile değiştirilebilir (ör. 646.01).

- **TL kalıntısı:** Döviz bakiyesi sıfır olduğu hâlde TL bakiyesi kalan hesapta TL bakiyesi kur farkıyla kapatılır.
- **Avanslar:** Sipariş avansı hesapları (159, 179, 340, 349, 440) parasal kalem sayılmadığından varsayılan olarak
  değerlemeye alınmaz. Dahil etmek için `--avanslari-dahil-et` kullanın.

| Kontrol | Önem |
|---|---|
| Kuru verilmemiş döviz | Yüksek |
| Döviz ve TL bakiyesinin yönü farklı | Yüksek |
| Beklenenin tersine bakiye (ör. alacak bakiyeli 120, borç bakiyeli 320) | Orta |
| Kayıtlı ortalama kurun değerleme kurundan %20'den fazla sapması (veri hatası olabilir) | Orta |
| TL kalıntısı | Orta |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 13 hesap, 31.12.2026
python main.py --bakiyeler bakiyeler.xlsx --kurlar kurlar.csv --tarih 31.12.2026
python main.py --bakiyeler b.xlsx --kurlar k.csv --kar-hesabi 646.01 --zarar-hesabi 656.01 --avanslari-dahil-et
```

| Dosya | Sütunlar |
|---|---|
| Bakiyeler | Hesap Kodu, Hesap Adı, Döviz, Döviz Bakiye, Borç/Alacak (B/A), TL Bakiye, Kur (isteğe bağlı, hesaba özel) |
| Kurlar | Döviz, Kur, Kaynak |

Borç/Alacak sütunu yoksa bakiyelerin işaretine bakılır: alacak bakiyesi eksi yazılır.

## Çıktı

`kur_degerleme.xlsx`:
- `Değerleme`: kayıtlı ortalama kur, değerleme kuru, değerlenmiş TL, kur farkı, durum; boş "Kontrol" sütunu.
- `Yevmiye Önerisi`: hesap bazında kur farkı satırları, 646 / 656 toplamları ve borç-alacak denge kontrolü.
- `Döviz Özeti`: döviz bazında net pozisyon, kur kaynağı ve kur farkı.
- `Uyarılar`.

## Dikkat

- **Kur seçimi:** Vergi Usul Kanunu md. 280'e göre yabancı paralar borsa rayici ile değerlenir. Borsada rayiç yoksa
  Hazine ve Maliye Bakanlığının belirlediği kur kullanılır. Uygulamada değerleme gününün TCMB döviz alış kuru
  kullanılır. Örnek kurlar gerçek değildir; değerleme tarihinin resmî kurlarını girin. TFRS / BOBİ FRS raporlaması
  için kullanılacak kur farklı olabilir.
- **Özel durumlar:**
  - Yatırım dönemine ait döviz kredisi kur farklarının ilgili varlığın maliyetine eklenmesi.
  - Avansların niteliği.
  - Enflasyon düzeltmesi etkileri.

  Bunlar ayrıca değerlendirilmelidir; mali müşavirinizle doğrulayın.
- **Kayıt önerisi:** Yevmiye satırları öneridir; muhasebe programına aktarmadan önce kontrol edin.
- **Örnek veri:** Örnek hesaplar ve kurlar kurgusaldır.

## Testler

Şunlar test edilir:
- Aktif ve pasif hesaplarda kur farkı; TL kalıntısı; kâr ve zarar toplamları.
- Avans hariç tutma ve dahil etme; hesaba özel kur.
- Kur yok, ters bakiye ve kur sapması uyarıları.
- Yevmiye dengesi, özel hesap kodları, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

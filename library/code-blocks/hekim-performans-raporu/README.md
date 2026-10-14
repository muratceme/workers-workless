# Hekim Performans Raporu · Kod Bloğu

> Hastane Yönetimi › Hastane Müdürü · Workers / Workless

Hekim bazında **muayene, ameliyat, yatış ve gelir** göstergelerini çalışma gününe böler; her hekimi **kendi bölümünün**
ortalamasıyla endeksler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Çalışma günü | Çalışma Günü sütunu; yoksa hekimin kaydı olan farklı gün sayısı |
| Günlük muayene / ameliyat / gelir | Toplam ÷ çalışma günü |
| Muayeneden yatış % | Yatış ÷ muayene × 100 |
| Bölüme göre endeks | Hekimin günlük değeri ÷ bölümün günlük değeri × 100 |
| Gelir payı % | Hekimin geliri ÷ bölüm geliri × 100 |

Bölümün günlük değeri, bölümdeki tüm hekimlerin toplamının toplam çalışma gününe bölümüdür. Dosyada her satır bir
hekimin bir dönemi ya da bir günü olabilir; hekim ve bölüm bazında toplanır.

**Bilgi notu:** bölümde birden çok hekim varsa endeksi 100 ± `--sapma` (varsayılan 50) dışında kalan göstergeler.

**Örnek:** Genel cerrahide bölüm günde 1,64 ameliyat yapıyor; Dr. E. B. 12 günde 31 ameliyat yapmış (günde 2,58) →
ameliyat endeksi **157,32**. Aynı hekimin muayene endeksi 63,10; az poliklinik, çok ameliyat yapan bir çalışma düzeni
olabilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                 # örnek veri
python main.py --veriler hekim_aylik_verileri.xlsx
python main.py --veriler hekim_aylik_verileri.xlsx --sapma 30
```

Zorunlu sütun: Hekim. Bölüm boşsa "Belirtilmemiş" sayılır.

## Çıktı

`cikti/hekim_performans_raporu.xlsx`:

- `Hekimler`: bölüm içinde gelire göre sıralı; toplamlar, günlük değerler, endeksler (sapanlar sarı).
- `Bölümler`: bölüm ve hastane toplamları.
- `Uyarılar`.

## Dikkat

- **Tek başına performans değerlendirmesi değildir.** Vaka karmaşıklığı, ameliyat grubu, nöbet, eğitim ve idari
  görevler, izin ve rapor günleri bu hesapta yoktur. Sayılar yalnız soru sormak için kullanılmalı; ücret, prim ya da
  disiplin kararı bu rapora dayandırılmamalıdır.
- **Kamu hastanelerindeki performans sistemi** ve özel hastanelerdeki sözleşme esasları ayrı kurallara tabidir; kod
  bunların puanlamasını uygulamaz.
- **Kişisel veri içerir.** Hekim adlarıyla rapor çalışan performans verisidir; erişimi sınırlayın.
- **Örnek veri:** Hekim baş harfleri ve sayılar kurgusaldır.

## Testler

- Çalışma günü (sütundan ya da tarihten), günlük değerler, muayeneden yatış, bölüm ortalaması, endeks, gelir payı.
- Sapma bilgi notu, tek hekimli bölüm, çalışma günü olmayan bölüm, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

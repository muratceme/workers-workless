# Resmî Yazı (Haciz/Müzekkere) Takibi · Kod Bloğu

> Bankacılık › Şube Bankacılığı › Operasyon Yetkilisi · Workers / Workless

İcra dairesi, vergi dairesi, SGK, mahkeme ve savcılıktan gelen **haciz ihbarnamesi, haciz bildirisi, haciz
kaldırma ve bilgi talebi** yazılarını tür, cevap süresi ve müşteri bazında takip eder. Süresi geçen ve
yaklaşanları uyarır. İnternete bağlanmaz.

## Ne yapar?

- **Tür tanıma:** Yazı türü serbest metinden tanınır:
  - "89/1", "birinci haciz ihbarnamesi" → İİK 89/1.
  - Benzer şekilde 89/2, 89/3, 6183 haciz bildirisi, haciz kaldırma (fekki), tedbir, bilgi/belge talebi.
- **Son cevap günü:**

  | Durum | Son gün |
  |---|---|
  | Yazıda süre yazıyorsa | Tebliğ + yazıdaki gün |
  | İİK 89/1 haciz ihbarnamesi | Tebliğ + 7 gün |
  | İİK 89/2 haciz ihbarnamesi | Tebliğ + 15 gün |
  | Haciz kaldırma | Tebliğ + `--fekk-gun` iş günü (banka içi hedef, varsayılan 1) |
  | Diğer türler (6183, 89/3, bilgi talebi…) | Yazıdaki süre girilmeli; girilmemişse "Süre girilmeli" |

  - Süreler takvim günüdür ve tebliği izleyen günden başlar.
  - Son gün hafta sonuna veya resmî tatile denk gelirse izleyen ilk iş gününe uzar.
  - Yazıdaki süre yasal süreden farklıysa yazıdaki esas alınır ve not düşülür.
- **Durum:** Süresi geçti (kritik), bugün son gün, yaklaşan (`--uyari-gun` iş günü), süresi var, süre girilmeli,
  cevaplandı, geç cevaplandı.
- **Zincir kontrolü:**
  - Aynı dosyada 89/1 süresinde cevaplanmadan 89/2 gelmişse kritik uyarı: menfi tespit süresi için hukuk
    birimine iletilmeli.
  - 89/3 her zaman kritik.
  - Haciz kaldırma yazısı gelen dosyalardaki açık haciz yazılarına not düşülür.
- **Veri kontrolleri:** TCKN/VKN kontrol hanesi, gelecekteki tebliğ tarihi, tebliğden önceki cevap tarihi,
  mükerrer kayıt.
- **Müşteri özeti:** Açık yazı sayısı ve toplam haciz tutarı (89/2 aynı alacak için tekrar sayılmaz), kurum
  sayısı ve en yakın son gün. Birden fazla kurumdan haciz gelen müşteriler işaretlenir; kredi izleme için de
  faydalıdır.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                      # örnek: 13 yazı, rapor tarihi 08.10.2026
python main.py --girdi gelen_yazilar.xlsx --tarih 09.10.2026
python main.py --girdi gelen_yazilar.xlsx --tatiller tatiller.txt --uyari-gun 3 --fekk-gun 2
```

| Sütun | Açıklama |
|---|---|
| **Tebliğ Tarihi**, **Yazı Türü** | Zorunlu |
| Kayıt No, Kurum, Dosya No, Müşteri No, Müşteri Adı, TCKN/VKN, Tutar | İsteğe bağlı |
| Yazıdaki Süre (Gün) | Yazıda süre belirtilmişse |
| Cevap Tarihi | Cevap/işlem yapıldıysa |
| Sorumlu | İsteğe bağlı |

**Tatiller:** Sabit genel tatiller (2912 sayılı Kanun) ve 2025-2027 dini bayramları kodda tanımlıdır. Dini
bayram tarihlerini ve idari izin günlerini her yıl kontrol edin; eksik günleri `--tatiller` ile verin (satır
başına `GG.AA.YYYY`).

## Çıktı

`resmi_yazi_takibi.xlsx`:
- `Özet`.
- `Takip Listesi`: duruma ve son güne göre sıralı, "Yapılan İşlem" sütunuyla.
- `Müşteri Özeti`.
- `Süre Kuralları`: kullanılan süreler ve dayanakları.

## Dikkat

- **Hukuki görüş değildir:** Bu bir takip aracıdır. Cevabın içeriği (bakiye bildirimi, itiraz, bloke) ve
  sürelerin yorumu için bankanızın hukuk biriminin görüşü ve yazıdaki ifade esastır.
- **Kodda tanımlı süreler:** Yalnız İİK 89/1 (7 gün) ve 89/2 (15 gün). Diğer türlerde süreyi yazıdan girin.
- **Tebliğ tarihi:** Elektronik tebligat ve UYAP üzerinden gelen yazılarda tebliğ tarihinin nasıl belirlendiğini
  kontrol edin.
- **Kişisel veri ve bankacılık sırrı:** Liste TCKN/VKN gibi kişisel veriler ve bankacılık sırrı içerir; raporu
  yetkisiz kişilerle paylaşmayın.
- **Örnek veri:** Örnek kurumlar, dosyalar ve kişiler kurgusaldır.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçlar kontrol edilmeden işlem yapılmamalıdır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

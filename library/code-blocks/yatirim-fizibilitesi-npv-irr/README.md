# Yatırım Fizibilitesi (NPV/IRR) · Kod Bloğu

> Finans › Finansal Analist · Workers / Workless

Yatırım nakit akışlarından net bugünkü değeri (NBD / NPV), iç verim oranını (İVO / IRR), değiştirilmiş İVO'yu,
geri dönüş süresini ve kârlılık endeksini hesaplar. İskonto oranına göre duyarlılık tablosu çıkarır. İnternete
bağlanmaz.

## Nasıl hesaplar?

- **Serbest nakit akışı** = gelir − nakit gider − vergi − yatırım harcaması − işletme sermayesi artışı + hurda değeri.
  Dosyada "Nakit Akışı" sütunu varsa doğrudan o kullanılır.
- **Vergi** = max(0, gelir − nakit gider − amortisman − devreden zarar) × vergi oranı. Zarar en fazla 5 yıl taşınır
  (KVK md. 9). Vergi oranı `--vergi` ile verilir; varsayılan %25'tir, güncel oranı kontrol edin.
- **NBD** = Σ NA_t ÷ (1 + r)^t. 0. yıl bugündür; akışlar yıl sonunda gerçekleşir.
- **Reel akış:** Akışlar sabit fiyatlarla hazırlandıysa `--akis reel --enflasyon 25` verin. Nominal iskonto oranı reel
  orana çevrilir: (1 + nominal) ÷ (1 + enflasyon) − 1. Nominal akışla reel oranı veya tersini karıştırmak sık
  yapılan bir hatadır.
- **İVO:** NBD'yi sıfırlayan oran, ikiye bölme yöntemiyle bulunur. Nakit akışının işareti birden çok kez değişiyorsa
  birden çok İVO olabilir; uyarı verilir.
- **Değiştirilmiş İVO (MIRR):** Negatif akışlar finansman oranıyla bugüne, pozitif akışlar yeniden yatırım oranıyla
  son yıla taşınır. Varsayılan iki oran da iskonto oranıdır.
- **Geri dönüş süresi:** Basit ve iskontolu kümülatif akışın pozitife döndüğü an; yıl içinde doğrusal varsayılır.
- **Kârlılık endeksi** = (NBD + bugünkü değerle yatırım) ÷ bugünkü değerle yatırım.
- **Duyarlılık:** İskonto oranı ±15 puan aralığında NBD tablosu ve grafiği.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 5 yıllık proje, %30 iskonto, %25 vergi
python main.py --akislar akislar.xlsx --iskonto 35 --vergi 25
python main.py --akislar akislar.xlsx --iskonto 35 --akis reel --enflasyon 25 --finansman 40 --yeniden-yatirim 30
```

| Sütun | Açıklama |
|---|---|
| Yıl | 0, 1, 2, … veya 2027, 2028, … (ardışık) |
| Yatırım Harcaması | Pozitif yazılır |
| Gelir, Nakit Gider | Amortisman nakit gidere dahil edilmez |
| Amortisman | Yalnız vergi hesabı için |
| İşletme Sermayesi Değişimi | Artış +, proje sonunda geri alınması − |
| Hurda Değeri | Vergi sonrası tutar olarak girin |
| Nakit Akışı | İsteğe bağlı; verilirse diğer sütunlar yerine kullanılır |

## Çıktı

`yatirim_fizibilitesi.xlsx`:
- `Özet`: NBD (renkli), İVO, MIRR, kârlılık endeksi, geri dönüş süreleri, değerlendirme.
- `Nakit Akışı`: yıllık vergi hesabı, zarar mahsubu, serbest nakit akışı, iskonto çarpanı, bugünkü değer, kümülatifler.
- `Duyarlılık`: iskonto oranına göre NBD ve grafik.
- `Varsayımlar`, `Uyarılar`.

## Dikkat

- **Girdiler sizindir:** Sonuç, satış tahminleri ve iskonto oranı (sermaye maliyeti) kadar güvenilirdir. Senaryo
  analizi için Senaryo ve Duyarlılık Analizi paketini kullanabilirsiniz.
- **Kapsam dışı:**
  - Yatırım teşvik belgesi kapsamındaki vergi indirimleri ve istisnalar
  - Finansman yapısı, hurda satış kazancının vergisi
  - Enflasyon düzeltmesi etkileri

  Bunları ayrıca değerlendirin.
- **Örnek veri:** Örnek proje kurgusaldır.

## Testler

Şunlar test edilir:
- Bilinen İVO örneği (−123.400; 36.200; 54.800; 48.100 → %5,96), NBD, MIRR ve geri dönüş formülleri.
- Vergi ve zarar mahsubu; örnek proje akışları.
- Reel / nominal dönüşümü; birden çok işaret değişimi; ardışık olmayan yıl hatası; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

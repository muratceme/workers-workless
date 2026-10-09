# KDV Beyannamesi Ön Kontrolü · Kod Bloğu

> Muhasebe › Genel Muhasebe Uzmanı · Workers / Workless

Satış ve alış fatura listelerini, mizanı ve hazırlanan KDV beyannamesi tutarlarını karşılaştırır. Uyumsuzlukları
beyanname verilmeden önce yakalar. İnternete bağlanmaz.

## Neleri kontrol eder?

**Fatura listeleri:**
- Dönem dışı tarih. Alışta önceki dönem faturası indirim süresi içindeyse kabul edilir.
- Mükerrer fatura; KDV ≠ matrah × oran.
- Oranın fatura tarihinde geçerli olması: 10.07.2023'ten itibaren %1, %10, %20.
- Tevkif edilen KDV = KDV × tevkifat oranı; istisna işlemde KDV olmaması.
- **İndirim süresi:** Alış faturasının ait olduğu takvim yılını izleyen yılın sonu geçmişse KDV indirilecek tutara
  alınmaz ve işaretlenir (KDVK md. 29/3).
- "İndirilebilir = Hayır" işaretli alışların KDV'si indirime alınmaz.

**Mutabakat (liste ↔ mizan ↔ beyanname):**

| Kalem | Liste | Mizan | Beyanname |
|---|---|---|---|
| Oran bazında matrah ve KDV (tevkifatsız teslimler) | ✓ | | ✓ |
| Kısmi tevkifat matrahı ve satıcının beyan ettiği KDV (KDV − tevkif edilen) | ✓ | | ✓ |
| İstisna matrahı | ✓ | | ✓ |
| Toplam hesaplanan KDV | ✓ | 391 | ✓ |
| Bu döneme ait indirilecek KDV | ✓ | 191 | ✓ |
| Önceki dönemden devreden KDV | | 190 | ✓ |
| Net satışlar ↔ satış matrahı (bilgi) | ✓ | 600–602 − 610–612 | |

**Beyanname aritmetiği:**
- Oran bazında matrah × oran = KDV.
- Toplam hesaplanan KDV = oran KDV'leri + kısmi tevkifat KDV'si.
- Ödenecek veya sonraki döneme devreden = hesaplanan − (devreden + indirilecek).

Alışlardaki tevkifat toplamı, 2 No'lu KDV beyannamesinde sorumlu sıfatıyla beyan edilecek tutar olarak ayrıca
gösterilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: Eylül 2026 (6 bilinçli hata)
python main.py --satislar s.xlsx --alislar a.xlsx --mizan mizan.xlsx --beyanname beyanname.csv --donem 2026-09
```

Beyanname dosyası iki sütunlu bir listedir (`Alan;Tutar`). Kullanılan alanlar:

| Alan | Açıklama |
|---|---|
| `Matrah %20`, `KDV %20`, `Matrah %10`, `KDV %10`, `Matrah %1`, `KDV %1` | Tevkifatsız teslim ve hizmetler |
| `Kısmi Tevkifat Matrahı`, `Kısmi Tevkifat Beyan Edilen KDV` | Kısmi tevkifatlı işlemler |
| `İstisna Matrahı` | İstisna kapsamındaki işlemler |
| `Toplam Hesaplanan KDV`, `Önceki Dönemden Devreden KDV`, `Bu Döneme Ait İndirilecek KDV` | |
| `Ödenecek KDV`, `Sonraki Döneme Devreden KDV` | |

Bu alanları beyanname taslağınızdaki karşılıklarından elle doldurun. Mizan dönem hareketleriyle verilir: Hesap Kodu,
Borç, Alacak. Alt hesaplı mizanda yaprak hesaplar toplanır.

## Çıktı

`kdv_beyanname_on_kontrol.xlsx`:
- `Mutabakat`: her kalem için liste, mizan ve beyanname tutarı; üç fark; durum.
- `Beyanname Aritmetiği`.
- `Liste Bulguları`: fatura bazında hatalar; boş "Düzeltme" sütunu.
- `Oran Özeti`: oran bazında satış ve alış; sorumlu sıfatıyla KDV.
- `Uyarılar`.

## Dikkat

- **Sade bir modeldir:** Beyannamenin tüm satırlarını kapsamaz. Ayrıca kontrol edilmesi gerekenler:
  - İade ve mahsup talepleri, ihraç kayıtlı ve tam tevkifatlı işlemler
  - Özel matrah şekilleri, indirimli orana tabi işlemler nedeniyle iade
  - İndirimler bölümündeki diğer satırlar
- **Mevzuat:** İndirim süresi, tevkifat kapsamı ve oranlar güncel KDV mevzuatına göre doğrulanmalıdır. Beyannameyi
  mali müşavir / sorumlu kişi onaylamalıdır.
- **Örnek veri:** Örnek faturalar, mizan ve beyanname kurgusaldır.

## Testler

Şunlar test edilir:
- Liste bulguları: dönem dışı, mükerrer, KDV hesabı, indirim süresi, indirilemez.
- Liste / mizan / beyanname mutabakatı; kısmi tevkifat; istisna.
- Beyanname aritmetiği hataları; dönem biçimleri; alt hesaplı mizan; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

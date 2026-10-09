# Fatura KDV Tutarlılık Kontrolü · Kod Bloğu

> Muhasebe › Muhasebe Elemanı · Workers / Workless

Fatura satırlarında ve toplamlarında matrah, KDV oranı, KDV tutarı ve tevkifat tutarlılığını kontrol eder.
Hesap hatalarını yuvarlama farklarından ayırır. İnternete bağlanmaz.

## Neleri kontrol eder?

| Seviye | Kontrol |
|---|---|
| Satır | Tutar = miktar × birim fiyat − iskonto |
| Satır | KDV oranı fatura tarihinde geçerli mi: 10.07.2023 ve sonrası %0, 1, 10, 20; öncesi %0, 1, 8, 18 |
| Satır | KDV = tutar × oran; oran %0 iken KDV olmamalı, oran > 0 iken KDV 0 olmamalı |
| Satır | Tevkifat oranı 2/10, 3/10, 4/10, 5/10, 7/10, 9/10 veya 10/10 mu; tevkif edilen KDV = KDV × tevkifat oranı |
| Satır | Aynı fatura + satıcı içinde tekrar eden sıra no; negatif / sıfır miktar (iade olabilir) |
| Fatura | Satırlar toplamı = mal/hizmet toplamı |
| Fatura | Toplam KDV: satır KDV'leri toplamı ve oran bazında matrah × oran ile karşılaştırılır |
| Fatura | Tevkifat toplamı; vergiler dahil = mal/hizmet + KDV; ödenecek = vergiler dahil − tevkifat |
| Fatura | Başlıkta olup satırı olmayan fatura |

**Fark sınıfları:**
- Tolerans (`--tolerans`, varsayılan 0,01 TL) içindeki fark tamam sayılır.
- Yuvarlama sınırına (`--yuvarlama`, 0,05 TL) kadar olan fark "Yuvarlama farkı" (bilgi) olarak raporlanır.
- Daha büyük fark hatadır.
- Fatura toplamlarında yuvarlama sınırı satır sayısı × 0,01 TL'ye kadar genişler.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 9 fatura, 17 satır (5 hata, 1 yuvarlama farkı)
python main.py --satirlar satirlar.xlsx --faturalar faturalar.xlsx
python main.py --satirlar e_fatura_listesi.xlsx    # E-Fatura Okuma ve Listeleme çıktısı (Faturalar + Satırlar sayfaları)
```

| Dosya | Sütunlar |
|---|---|
| Satırlar | Fatura No, Tarih, Satıcı, Sıra, Mal/Hizmet, Miktar, Birim Fiyat, İskonto, Tutar, KDV Oranı, KDV Tutarı, Tevkifat Oranı (9/10, %90 veya 0,9), Tevkif Edilen KDV |
| Faturalar (isteğe bağlı) | Fatura No, Satıcı, Mal/Hizmet Toplamı, Toplam KDV, KDV Tevkifatı, Vergiler Dahil, Ödenecek |

## Çıktı

`fatura_kdv_kontrolu.xlsx`:
- `Bulgular`: önem, tür, fatura, satır, açıklama (beklenen ve beyan edilen değer); boş "Karar / Düzeltme" sütunu.
- `Fatura Özeti`: satır ve fatura toplamları yan yana; durum (Tamam / Yuvarlama / Hata).
- `Satırlar`: beyan edilen ve hesaplanan tutar, KDV ve tevkifat.
- `KDV Oran Özeti`: oran bazında matrah, KDV ve matrah × oran farkı.

## Dikkat

- **Ürün oranı kontrol edilmez:** Bir mal veya hizmete hangi KDV oranının uygulanacağı (I) ve (II) sayılı
  listelere ve istisnalara bağlıdır. Paket yalnız oranın o tarihte var olan bir oran olup olmadığına bakar.
- **Tevkifat kapsamı:** İşlemin tevkifata tabi olup olmadığı, hangi oranın uygulanacağı, tevkifat kodu ve tutar sınırı
  KDV Genel Uygulama Tebliği'ne göre belirlenir. Bunları güncel mevzuatla ve mali müşavirinizle doğrulayın; paket
  yalnız oranın tanımlı oranlardan biri olup olmadığını ve hesabı kontrol eder.
- **Döviz:** Döviz cinsinden faturalarda tutarlar fatura para birimiyle karşılaştırılır.
- **Örnek veri:** Örnek faturalar kurgusaldır.

## Testler

Şunlar test edilir:
- Örnekteki 5 hata ve yuvarlama farkı; tarihe göre eski oranlar (%18).
- Tevkifat oranı biçimleri; mükerrer satır; fatura toplamları.
- E-fatura paketi çıktısının (iki sayfalı Excel) okunması; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

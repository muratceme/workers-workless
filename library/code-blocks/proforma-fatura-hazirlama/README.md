# Proforma Fatura Hazırlama · Kod Bloğu

> Dış Ticaret › Dış Ticaret Uzmanı · Workers / Workless

Sipariş kalemleri, ürün kartları ve teslim şekli bilgisinden İngilizce **proforma fatura** ve **çeki listesi**
(packing list) hazırlar. Teslim şeklini Incoterms® 2020 kurallarına göre kontrol eder; GTİP, menşe, ağırlık ve banka
bilgilerindeki eksikleri ayrı bir iç kontrol dosyasında listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Satırlar:** Tutar = miktar × birim fiyat. Siparişte "Birim Fiyat" yazılıysa o, yazılı değilse ürün kartındaki
  fiyat kullanılır.
  - Ürün kartındaki döviz proforma dövizinden farklıysa ve siparişte fiyat yoksa satır alınmaz. Paket döviz çevrimi
    yapmaz.
- **Toplam:** Mal bedeli + navlun + sigorta. Navlun ve sigorta yalnız teslim şekli gerektiriyorsa eklenir. Toplamın
  İngilizce yazıyla karşılığı da yazılır ("SAY EUR … ONLY").
- **Çeki listesi:**
  - koli sayısı = ⌈miktar / koli içi adet⌉;
  - net ağırlık = miktar × birim net ağırlık;
  - brüt ağırlık = net + koli sayısı × koli darası;
  - hacim = koli sayısı × koli ölçüleri (m³);
  - koli numaraları sıralı verilir (ör. 331–381).

### Incoterms® 2020 kontrolleri

| Kural | Taşıma türü | Navlun satıcıda | Sigorta satıcıda |
|---|---|---|---|
| EXW, FCA | Her tür | Hayır | Hayır |
| CPT, DAP, DPU, DDP | Her tür | Evet | Hayır |
| CIP | Her tür | Evet | Evet |
| FAS, FOB | Yalnız deniz / iç su yolu | Hayır | Hayır |
| CFR | Yalnız deniz / iç su yolu | Evet | Hayır |
| CIF | Yalnız deniz / iç su yolu | Evet | Evet |

- Deniz kuralı (FAS, FOB, CFR, CIF) deniz dışı taşımayla yazılmışsa uyarı verilir ve karşılığı önerilir (FCA, CPT,
  CIP).
- Teslim yeri zorunludur: "CIF Hamburg Port, Incoterms® 2020".
- CIF ve CIP'te satıcı sigorta yaptırır: en az mal bedelinin %110'u; CIF için ICC (C), CIP için ICC (A) teminatı.
- Gerekmeyen navlun veya sigorta tutarı girilmişse proformaya eklenmez ve uyarı verilir.
- DAT yazılmışsa DPU önerilir (Incoterms 2020 değişikliği).

| Uyarı | Önem |
|---|---|
| Ürün kartı yok, miktar veya fiyat yok, döviz uyuşmuyor (kalem alınmaz) | Yüksek |
| Teslim şekli yok / bilinmiyor / taşıma türüne uymuyor; teslim yeri yok; alıcı yok | Yüksek |
| GTİP 6–12 hane değil; menşe, tanım, ağırlık veya koli bilgisi yok | Orta |
| Navlun veya sigorta eksik / gereksiz; ödeme şekli yok; banka bilgisi eksik | Orta |
| Kısmi koli; liste fiyatının altında anlaşılan fiyat | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek: 6 kalem, CIF Hamburg, EUR
python main.py --siparis siparis.xlsx --urunler urunler.xlsx --bilgi proforma_bilgisi.csv --cikti proformalar
```

| Dosya | Sütunlar |
|---|---|
| Sipariş | Ürün Kodu, Miktar, Birim Fiyat (anlaşılan fiyat; isteğe bağlı), Not |
| Ürün kartları | Ürün Kodu, Tanım (İngilizce), GTİP, Menşe, Birim (PCS, SET, KG...), Birim Fiyat, Döviz, Net Ağırlık (kg), Koli İçi Adet, Koli Darası (kg), Koli Ölçüleri (cm; 60x40x45) |
| Proforma bilgisi | İki sütun (Alan;Değer): Proforma No, Tarih, Geçerlilik (gün), Satıcı, Satıcı Adresi, Alıcı, Alıcı Adresi, Ülke, Teslim Şekli, Teslim Yeri, Taşıma Şekli (Deniz / Karayolu / Havayolu), Yükleme Limanı, Varış Limanı, Tahmini Sevk, Ödeme Şekli, Döviz, Navlun, Sigorta, Banka, IBAN, SWIFT, Notlar |

Alıcıya gidecek alanları (satıcı, alıcı, ödeme şekli, notlar) İngilizce yazın; paket çeviri yapmaz.

## Çıktı

- `proforma_<no>.xlsx` alıcıya gönderilir:
  - `Proforma Invoice`: taraflar, teslim şekli, satırlar (GTİP ve menşe ile), toplam, yazıyla tutar, ödeme şekli,
    geçerlilik, paketleme özeti, banka bilgileri.
  - `Packing List`: koli sayısı ve numaraları, net / brüt ağırlık, hacim.
- `proforma_<no>_kontrol.xlsx` şirket içinde kalır: özet, "Gönderime hazır mı?" değerlendirmesi ve uyarılar.

## Dikkat

- **Proforma ticari fatura değildir.** İhracat için e-fatura / e-arşiv düzenlenmesi ve gümrük beyannamesi ayrı
  işlemlerdir.
- **Incoterms® kuralları** yalnız teslim, risk ve masraf paylaşımını düzenler; mülkiyet devri ve ödeme koşulları
  sözleşmede ayrıca belirlenir. Kuralların resmî metni ICC yayınıdır.
- **GTİP ve menşe** yalnız biçim olarak kontrol edilir. Doğru tarife pozisyonu ve menşe kuralları için gümrük
  müşavirinize danışın.
- **Ağırlık ve ölçüler** ürün kartındaki değerlerden hesaplanır. Sevkiyat öncesi fiilî tartı ve ölçümle doğrulayın.
- **Örnek veri:** Örnek firmalar, banka bilgileri ve fiyatlar kurgusaldır.

## Testler

Şunlar test edilir:
- Mal bedeli, navlun, toplam ve yazıyla tutar (elle hesaplanmış).
- Koli sayısı, koli numaraları, net / brüt ağırlık ve hacim; kısmi koli.
- Incoterms kontrolleri: FOB'da gereksiz navlun, deniz kuralı ile karayolu, DAT, CIP'te sigorta, teslim yeri.
- Döviz uyuşmazlığı, eksik ürün kartı ve menşe; alıcı dosyasında iç uyarı olmaması; CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

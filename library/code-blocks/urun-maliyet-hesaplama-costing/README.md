# Ürün Maliyet Hesaplama (Costing) · Kod Bloğu

> Tekstil ve Konfeksiyon › Müşteri Temsilciliği (Merchandising) › Development Merchandiser · Workers / Workless

Hazır giyim için maliyet föyü (cost sheet) hazırlar ve FOB teklif fiyatını hesaplar. Föy kumaş, aksesuar, fason
işlem, CM (kesim-dikim), genel gider ve sipariş başı sabit giderlerden oluşur. Farklı para birimlerindeki kalemleri
teklif para birimine çevirir. Müşteri hedef fiyatıyla karşılaştırır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kalem | Hesap (adet başı) |
|---|---|
| Kumaş | tüketim (kg veya m) × (1 + fire %) × birim fiyat |
| Aksesuar, fason işlem (baskı, nakış, yıkama) | miktar × (1 + fire %) × birim fiyat |
| CM | SAM (dk) × dakika maliyeti ÷ hat verimliliği |
| Diğer | miktar × birim fiyat (ör. limana nakliye) |
| Sabit | sipariş başı tutar ÷ sipariş adedi (test raporu, numune maliyeti) |
| Genel gider | üretim maliyeti × genel gider % |
| **FOB teklif** | **toplam maliyet ÷ (1 − kâr marjı % − komisyon %)** |

**Ayrıntılar:**
- **Marj ve komisyon:** İkisi de satış fiyatı üzerinden hesaplanır. Örneğin %12 marj ve %3 komisyonla maliyet
  0,85'e bölünür. Maliyet üzerine kâr (markup) kullanıyorsanız oranı buna göre çevirin.
- **Hedef fiyat:** Hedef fiyat verilirse o fiyattaki marj hesaplanır. Ayrıca istenen marjı korumak için adet başı ne
  kadar maliyet düşüşü gerektiği gösterilir.
- **Duyarlılık:**
  - Teklif dövizinin TL kuru ±%5 değişirse: TL maliyetler dövize daha pahalı ya da daha ucuz çevrilir.
  - Kumaş fiyatları ±%10 değişirse.
- **Uyarılar:**
  - Fiyatı veya kuru eksik kalemler hesaba katılmaz ve "Yüksek" uyarıyla raporlanır.
  - Kumaş ya da CM kalemi olmayan model uyarılır.
  - Firesi 0 olan kumaş uyarılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek: 3 model, örnek kurlar
python main.py --modeller modeller.xlsx --kalemler maliyet_kalemleri.xlsx --kur USD=41,20 EUR=48,10
python main.py --modeller m.xlsx --kalemler k.xlsx --kur USD=41,20 EUR=48,10 GBP=55,30
```

| Dosya | Sütunlar |
|---|---|
| Modeller | Model, Müşteri, Açıklama, Sipariş Adedi, Teklif Para Birimi, Verimlilik %, Genel Gider %, Kâr Marjı %, Komisyon %, Hedef Fiyat |
| Maliyet kalemleri | Model, Grup, Kalem, Miktar, Birim, Fire %, Birim Fiyat, Para Birimi (TL / USD / EUR / GBP) |

**Grup ve miktar:**
- **Grup:** Kumaş, Aksesuar, Fason İşlem, CM, Diğer veya Sabit olabilir.
- **CM satırları:** Miktar SAM'dır (dk), birim fiyat dakika maliyetidir. Kesim, dikim ve ütü-paket ayrı satır
  yazılabilir.
- **Yüzdeler:** "8" veya "0,08" yazılabilir.

## Çıktı

`maliyet_foyu.xlsx`:
- `Özet`: model bazında grup toplamları, üretim maliyeti, genel gider, toplam maliyet, FOB teklif, hedef fiyattaki
  marj, gereken düşüş ve boş "Karar / Not" sütunu.
- `Föy <model>`: her model için kalem kalem maliyet föyü; adet başı tutar, maliyetteki pay ve grup ara toplamları.
- `Duyarlılık`: kur ve kumaş senaryolarında toplam maliyet ve FOB.
- `Uyarılar`.

## Dikkat

- **Örnek değerler:** Örnek veri yalnız gösterim içindir; gerçek piyasa fiyatı değildir. Bu değerler şunlardır:
  - kurlar (USD = 41,20, EUR = 48,10 TL)
  - dakika maliyeti (4,10 TL)
  - verimlilik
  - kumaş ve aksesuar fiyatları
  - genel gider, marj ve komisyon oranları
- **Dakika maliyeti:** Kendi işletmenizin maliyetinden hesaplanmalıdır: aylık işçilik ve üretim giderleri ÷ çalışılan
  dakika. Fason dikimde atölyenin adet fiyatını "Fason İşlem" grubuna yazabilirsiniz.
- **Tüketim:** Kumaş tüketimi pastal (marker) veya modelhane verisinden alınmalıdır. Beden dağılımına göre ortalama
  tüketim kullanın.
- **Kapsam dışı:** Navlun, sigorta, finansman maliyeti ve ihracat teşvikleri föye eklenmedikçe hesaba katılmaz. FOB
  dışı teslim şekillerinde (CIF, DDP) ilgili giderleri "Diğer" veya "Sabit" grubuna ekleyin.
- **Örnek veri:** Örnek müşteriler ve modeller kurgusaldır.

## Testler

Şunlar test edilir:
- Elle hesaplanmış föy değerleri (kumaş, CM, sabit, FOB, hedef marjı).
- Kur çevirisi ve duyarlılık.
- Eksik fiyat / kur uyarıları, yüzde biçimleri, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

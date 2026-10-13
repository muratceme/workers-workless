# Araç Satış Teklifi Hesaplama · Kod Bloğu

> Satış › Satış Danışmanı · Workers / Workless

Liste fiyatı, kampanya, takas, peşinat, kredi ve sigorta kalemlerinden müşteriye sunulacak **aylık taksiti, ilk gün
ödenecek tutarı ve toplam maliyeti** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Net araç fiyatı | Liste fiyatı − kampanya indirimi |
| Kredi tutarı | Net araç fiyatı − takas değeri − peşinat |
| Aylık maliyet oranı | Aylık faiz × (1 + (KKDF + BSMV) ÷ 100) |
| Aylık taksit | Kredi × oran ÷ (1 − (1 + oran)^−vade); oran 0 ise kredi ÷ vade |
| İlk gün ödenecek | Peşinat + kasko + trafik sigortası + diğer masraflar (kredi yoksa kalan tutar da) |
| Toplam maliyet | Net araç fiyatı + kredi faizi ve vergisi + sigorta ve masraflar |

Ödeme planında her ay faiz kalan anapara üzerinden, vergi faiz üzerinden hesaplanır; son taksit kalan anaparayı
kapatır (kuruş farkı son taksite yansır).

**Örnek:** Liste 1.000.000, indirim 50.000 → net 950.000. Takas 300.000, peşinat 250.000 → kredi 400.000. 12 ay,
aylık %3 → taksit 40.184,83. Kasko, trafik ve masraf 45.000 → ilk gün ödenecek 295.000.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --teklifler teklifler.xlsx
python main.py --teklifler teklifler.xlsx --kkdf 15 --bsmv 15
```

Zorunlu sütun: Liste Fiyatı. Diğerleri: Teklif No, Müşteri, Araç, Kampanya İndirimi, Takas Değeri, Peşinat, Kredi
Vadesi (ay), Aylık Faiz %, Kasko, Trafik Sigortası, Diğer Masraflar. Boş bırakılan tutar 0 sayılır.

## Çıktı

`cikti/arac_satis_teklifi.xlsx`:

- `Teklifler`: her teklif için net fiyat, kredi, taksit, toplam geri ödeme, ilk gün ödenecek ve toplam maliyet.
- `Ödeme Planı`: kredili tekliflerin aylık taksit, faiz, vergi ve anapara dökümü.
- `Uyarılar`: vadesi girilmemiş kalan tutar, araç fiyatını aşan takas ve peşinat, şüpheli faiz oranı.

## Dikkat

- **Bağlayıcı kredi teklifi değildir.** Faiz oranını, KKDF ve BSMV'yi siz girersiniz; varsayılan 0'dır. Kesin ödeme
  planını, tahsis ücretini ve sigorta şartlarını banka belirler.
- **Taşıt kredisi sınırları denetlenmez.** Araç bedeline göre azami kredilendirme oranı ve azami vade düzenlemeyle
  belirlenir ve değişir; teklif vermeden önce güncel sınırları kontrol edin.
- **Faiz aylık girilir.** Yıllık oran yazarsanız taksit çok yüksek çıkar; %10 üzeri aylık oranlar uyarılır.
- **Liste fiyatının neleri içerdiğini siz belirlersiniz.** ÖTV, KDV, plaka ve tescil dahil "anahtar teslim" fiyat
  kullanmanız karşılaştırmayı kolaylaştırır; kod vergi hesaplamaz.
- **Takas değeri ekspertiz sonucudur.** Kod değerleme yapmaz.
- **Örnek veri:** Müşteriler, araçlar, fiyatlar ve oranlar kurgusaldır.

## Testler

- Eşit taksit formülü, ödeme planı, son taksit düzeltmesi, sıfır faiz, KKDF ve BSMV etkisi.
- Kredili, nakit ve peşin teklifler; araç fiyatını aşan takas ve peşinat; uyarılar.
- Okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

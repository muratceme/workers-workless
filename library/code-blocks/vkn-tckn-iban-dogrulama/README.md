# VKN / TCKN / IBAN Doğrulama · Kod Bloğu

> Muhasebe › Ön Muhasebe Elemanı · Bankacılık › Gişe Yetkilisi · Workers / Workless

Cari kart, personel veya ödeme listelerindeki **vergi kimlik numaralarını (VKN)**, **T.C. kimlik
numaralarını (TCKN)** ve **IBAN'ları** toplu olarak doğrular. Hatalı, eksik haneli ve mükerrer
kayıtları renklendirilmiş bir Excel raporunda gösterir. İnternete bağlanmaz.

> Bu araç numaranın **matematiksel olarak geçerli** olup olmadığını kontrol eder; numaranın
> gerçekten o kişiye/firmaya ait olduğunu doğrulamaz. Bunun için GİB ve NVİ sorgu hizmetlerini kullanın.

## Ne kontrol eder?

| Tür | Kural |
|---|---|
| VKN | 10 hane; son hane kontrol hanesi (GİB algoritması) |
| TCKN | 11 hane, ilk hane 0 değil; 10. ve 11. haneler kontrol haneleri |
| IBAN | ISO 13616 mod 97 kontrolü; TR IBAN için 26 karakter, rezerv alan (10. karakter) 0 |
| Ortak | Boşluk, nokta, tire temizliği; aynı numaranın listede birden çok geçmesi (mükerrer) |
| Excel | Sayı olarak saklandığı için baştaki 0'ı silinmiş 9 haneli VKN'ler tespit edilip düzeltilir |

"Vergi No" sütununda 10 hane VKN, 11 hane TCKN (şahıs işletmeleri) olarak ayrı ayrı doğrulanır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                        # örnek veriyle dener
python main.py --girdi cari_kartlar.xlsx               # sütunları otomatik bulur
python main.py --girdi liste.csv --sutun "Müşteri Vergi No" --sutun "Banka Hesabı"
```

## Girdi

İlk satırı başlık olan `.xlsx` veya `.csv` (`,` `;` veya sekme ayraçlı). Şu başlıklar otomatik tanınır:
`VKN`, `Vergi No`, `Vergi Kimlik No`, `TCKN`, `TC Kimlik No`, `Kimlik No`, `IBAN`.

## Çıktı (`cikti/dogrulama.xlsx`)

- **Doğrulama:** orijinal tablo + her kontrol edilen sütun için Tür, Normal, Durum, Açıklama.
  Hatalı satırlar kırmızı, yalnızca uyarı (mükerrer) olanlar sarı.
- **Hatalılar:** Excel satır numarasıyla hatalı değerler ve nedeni.
- **Özet:** sütun, tür ve durum bazında adetler.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

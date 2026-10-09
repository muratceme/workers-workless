# Cari Hesap Ekstresi Hazırlama · Kod Bloğu

> Muhasebe › Ön Muhasebe Elemanı · Workers / Workless

Muhasebe hareketlerinden müşteri ve tedarikçi bazında dönem ekstresi ve bakiye özeti hazırlar. Açık kalemleri ve
vadesi geçmiş tutarları çıkarır, her cari için mutabakat yazısı taslağı üretir. İnternete bağlanmaz.

## Nasıl çalışır?

- **Ekstre:** Dönem başından önceki hareketler devreden bakiye olur. Dönem hareketleri tarih sırasıyla yürüyen
  bakiyeyle listelenir; kapanış bakiyesi B/A olarak gösterilir.
- **Açık kalemler (FIFO):** Bakiyenin yönündeki kalemler (müşteride satış faturaları, tedarikçide alış faturaları)
  ödemelerle en eskiden başlayarak kapatılır. Kalan kalemlerin vadesi dönem sonundan önceyse "vadesi geçmiş" sayılır.
- **Vade:** Vade tarihi yoksa cari kartındaki vade günü (yoksa `--vade`) belge tarihine eklenir. Tahsilat, ödeme,
  iade ve çek kayıtlarına vade yazılmaz.
- **Müşteri / tedarikçi ayrımı:** Cari listesindeki "Tür" sütunundan, yoksa hesap kodundan (120 / 320) anlaşılır.

| Kontrol | Önem |
|---|---|
| Aynı belge no ve tutarla tekrar eden kayıt | Yüksek |
| Ters bakiye: alacak bakiyeli müşteri, borç bakiyeli tedarikçi (avans / fazla ödeme) | Orta |
| Dönemde hareketi olmayan bakiye | Orta |
| Vadesi geçmiş tutar ve en eski vade | Bilgi |
| Cari listesinde olmayan hareket | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 6 cari, 01.07–30.09.2026
python main.py --hareketler hareketler.xlsx --cariler cariler.xlsx --donem 01.07.2026 30.09.2026
python main.py --hareketler h.xlsx --donem 01.01.2026 30.09.2026 --ayri-dosya --firma "Örnek AŞ" --cevap-suresi 10
```

| Dosya | Sütunlar |
|---|---|
| Hareketler | Tarih, Cari Kod, Cari Unvan (isteğe bağlı), Belge Türü, Belge No, Açıklama, Borç, Alacak, Vade Tarihi |
| Cariler (isteğe bağlı) | Cari Kod, Unvan, Tür (Müşteri / Tedarikçi), Vade (gün) |

Muhasebe programınızın "cari hesap hareketleri" veya "muavin defter" raporunu Excel'e aktarıp kullanabilirsiniz.

## Çıktı

`cari_ekstreler.xlsx`:
- `Bakiye Özeti`: devir, dönem borç ve alacak, bakiye, açık kalem sayısı, vadesi geçmiş tutar, en eski vade; boş
  "Mutabakat Durumu" sütunu.
- `Ekstreler`: cari başına blok; devir, hareketler, yürüyen bakiye, kapanış.
- `Açık Kalemler`: FIFO sonrası açık tutar ve gecikme günü.
- `Mutabakat Metinleri`: cari başına yazı taslağı; boş "Gönderildi" ve "Cevap" sütunları.
- `Uyarılar`.

`--ayri-dosya` ile her cari için ekstre ve mutabakat metni içeren ayrı bir Excel dosyası da üretilir
(`cari_ekstreler_cari/` klasörü).

## Dikkat

- **FIFO varsayımdır:** Ödemenin hangi faturaya karşılık yapıldığı kayıtlarda belirtilmişse sonuç farklı olabilir.
  Çek ve senetle yapılan tahsilatlar, tahsil edilmemiş olsa da kayıt tarihinde kapatıcı sayılır.
- **Mutabakat metni taslaktır:** Göndermeden önce bakiyeyi ve alıcı bilgisini kontrol edin.
- **Kişisel veri:** Şahıs firmalarında unvan kişisel veridir. Ayrı ekstre dosyalarını yalnız ilgili cariye gönderin.
- **Örnek veri:** Örnek cariler ve hareketler kurgusaldır.

## Testler

Şunlar test edilir:
- Devir, dönem ve kapanış bakiyeleri (müşteri ve tedarikçi).
- FIFO açık kalemler; cari vadesi ve varsayılan vade; ödemelere vade yazılmaması.
- Mükerrer, ters bakiye, hareketsiz bakiye ve vadesi geçmiş uyarıları.
- Mutabakat metni, ayrı dosyalar, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

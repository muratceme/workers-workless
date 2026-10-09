# Satın Alma Talebi Konsolidasyonu · Kod Bloğu

> Satın Alma › Satın Alma Uzman Yardımcısı · Workers / Workless

Departmanlardan gelen satın alma taleplerini malzeme bazında birleştirir. Stok ve açık siparişleri düşerek net
ihtiyacı bulur ve tedarikçilere gönderilecek **teklif istek listesini** çıkarır. İstenirse her tedarikçi için ayrı bir
teklif formu yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Onay:** Yalnız onaylı talepler birleştirilir. Onay bekleyen ve reddedilen talepler ayrı sayfada listelenir. Onay
  sütunu yoksa tüm talepler onaylı sayılır.
- **Malzeme anahtarı:** Malzeme kodu; kod yoksa malzeme adı. Ad karşılaştırılırken büyük-küçük harf ve Türkçe
  karakter farkı yok sayılır.
- **Ölçü birimi:** Eş adlar birleştirilir (ad / adet / tane, kilogram / kg, litre / lt...). Aynı malzeme farklı
  birimle istenmişse (ör. koli ve paket) satırlar birleştirilmez, uyarı verilir. Kart birimi farklı olan satırda stok,
  asgari sipariş ve fiyat uygulanmaz.
- **Net ihtiyaç:** talep − kullanılabilir stok − açık sipariş (en az 0).
  - Kullanılabilir stok = eldeki stok − emniyet stoğu (en az 0).
  - Stok dosyası verilmezse net ihtiyaç talep toplamına eşittir.
- **Sipariş miktarı:** Net ihtiyaç asgari sipariş miktarına çıkarılır, sonra sipariş katına yukarı yuvarlanır.
  Örnek: 330 çift, kat 50 → 350 çift.
- **Tahmini tutar:** sipariş miktarı × son alış fiyatı (malzeme kartından).
- **Gerekli teklif:** Tahmini tutar `--uc-teklif-esik` (varsayılan 100.000 TL) ve üzerindeyse en az 3 teklif, altındaysa
  1. Bu eşik **şirket satın alma yönetmeliğine** göre değiştirilmelidir.
- **Teklif istek listesi:** Sipariş gereken her malzeme, kartındaki her tedarikçi için bir satır olur. Tedarikçisi
  olmayan malzeme "(Belirlenecek)" olarak listelenir. Teklif son tarihi = rapor tarihi + `--teklif-gun` (7).

| Kontrol | Önem |
|---|---|
| Miktarı okunamayan veya sıfır talep (alınmaz) | Yüksek |
| İstenen teslim tarihi geçmiş | Yüksek |
| Termin riski: rapor tarihi + teklif süresi + tedarik süresi > istenen teslim | Yüksek |
| Olası mükerrer talep: aynı departman ve malzeme, `--mukerrer-gun` (7) içinde iki talep (ikisi de toplanır) | Orta |
| Kodsuz talepte olası aynı malzeme (ad benzerliği; birleştirilmez) | Orta |
| Aynı malzeme farklı birimle istenmiş | Orta |
| 3 teklif gereken malzemede kayıtlı tedarikçi 3'ten az | Orta |
| Son alış fiyatı bilinmiyor | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                 # örnek: 25 talep satırı, 7 departman, rapor tarihi 09.10.2026
python main.py --talepler talepler.xlsx --malzemeler malzemeler.xlsx --stok stok.xlsx --bugun 09.10.2026
python main.py --talepler talepler.xlsx --malzemeler malzemeler.xlsx --ayri-dosya --uc-teklif-esik 250000
```

| Dosya | Sütunlar |
|---|---|
| Talepler | Talep No, Talep Tarihi, Departman, Talep Eden, Malzeme Kodu, Malzeme Adı, Miktar, Birim, İstenen Teslim, Onay Durumu (Onaylandı / Onay bekliyor / Reddedildi), Gerekçe |
| Malzeme kartı (isteğe bağlı) | Malzeme Kodu, Malzeme Adı, Kategori, Tedarikçiler (`/` ile ayrılmış), Son Alış Fiyatı, Döviz, Birim, Asgari Sipariş, Sipariş Katı, Tedarik Süresi (gün) |
| Stok (isteğe bağlı) | Malzeme Kodu, Eldeki Stok, Emniyet Stoğu, Açık Sipariş |

## Çıktı

`satin_alma_konsolidasyonu.xlsx`:
- `Özet`.
- `Konsolide Liste`: talep toplamı, stok, açık sipariş, net ihtiyaç, sipariş miktarı, tahmini tutar, gerekli teklif
  sayısı, departman dağılımı ve talep numaraları.
- `Teklif İstek Listesi`: tedarikçi × malzeme; boş "Gönderildi mi?" sütunuyla.
- `Bekleyen Talepler`: onay bekleyen ve reddedilen talepler.
- `Talep Detayı`: tüm satırlar ve kontrol notları.
- `Uyarılar`.

`--ayri-dosya` ile `teklif_formlari/<tedarikçi>.xlsx`: tedarikçinin dolduracağı birim fiyat, para birimi, teslim
süresi ve teklif geçerlilik alanları sarı.

## Dikkat

- **Teklif eşiği ve teklif sayısı** şirket politikasıdır. Kamu alımları 4734 sayılı Kamu İhale Kanunu'na tabidir; bu
  paket kamu ihale süreci için tasarlanmamıştır.
- **Stok verisi** rapor anındaki ERP stokuyla aynı olmalıdır. Rezerve edilmiş stok "eldeki" stoktan önce düşülmelidir.
- **Birim dönüşümü** yapılmaz. Koli / paket gibi farklı birimleri tek birime çevirip yeniden çalıştırın.
- **Ad benzerliği** yalnız uyarıdır. Aynı malzemeyi malzeme koduna bağlamak kullanıcıya bırakılmıştır.
- **Örnek veri:** Örnek talepler, malzemeler, fiyatlar ve tedarikçi adları kurgusaldır.

## Testler

Şunlar test edilir:
- Onay durumları ve geçersiz miktar.
- Net ihtiyaç: emniyet stoğu, açık sipariş, asgari sipariş, sipariş katı ve stoktan karşılanan satır.
- Birim uyuşmazlığı.
- Mükerrer talep, ad benzerliği, geçmiş teslim, termin riski, tedarikçi havuzu kontrolleri.
- Teklif listesi, ayrı teklif formları, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

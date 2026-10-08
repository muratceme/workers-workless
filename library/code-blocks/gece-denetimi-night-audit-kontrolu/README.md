# Gece Denetimi (Night Audit) Kontrolü · Kod Bloğu

> Turizm ve Otelcilik › Ön Büro › Night Auditor (Gece Denetçisi) · Workers / Workless

Gün kapatılmadan (date roll) önce PMS'ten alınan raporları karşılaştırır. Açık folyoları, fiyat hatalarını, oda
durumu uyuşmazlıklarını ve kasa farklarını tek listede toplar. İnternete ve PMS'e bağlanmaz; PMS'ten dışa aktarılan
Excel/CSV dosyalarıyla çalışır.

## Kontroller

| Alan | Kontrol | Seviye |
|---|---|---|
| Folyo | Çıkış yapmış ama bakiyesi kapanmamış (açık) folyo | Hata |
| Folyo | Çıkış tarihi geçmiş ama hâlâ konaklıyor · bugün çıkışı olup çıkış yapılmamış | Hata · Dikkat |
| Folyo | Bakiye limiti aşımı (`--bakiye-limiti`), no-show bakiyesi, mükerrer folyo, aynı odada birden fazla folyo | Dikkat / Hata |
| Fiyat | Konaklayan misafire oda ücreti postalanmamış | Hata |
| Fiyat | Postalanan ücret ↔ fiyat kodu + oda tipine göre fiyat listesi | Yüksek |
| Fiyat | Ücretsiz oda (COMP / house use): onay kontrolü | Dikkat |
| Oda | **Skip:** PMS dolu, kat hizmetleri boş · **Sleep:** PMS boş, kat hizmetleri dolu · arızalı (OOO) odada misafir | Yüksek / Hata |
| Kasa | Kasiyer × ödeme tipi: sistem tahsilatı ↔ kasa sayımı / POS gün sonu | Hata |
| Kasa | Folyo listesinde olmayan folyoya tahsilat | Dikkat |
| Gelir | Folyolardan hesaplanan oda geliri ↔ gelir raporundaki oda geliri (`--rapor-oda-geliri`) | Hata |

**Özet sayfası:** doluluk, oda geliri, ADR ve RevPAR. Satılabilir oda sayısına OOO odalar dahil edilmez; ADR
ücretli oda başına hesaplanır.

**Durum değerleri:** Durum sütunu Türkçe veya İngilizce olabilir (Konaklıyor / In-House, Çıkış Yaptı / Checked Out,
No-Show). Kat hizmetleri durumu için Dolu / Boş / OOO ya da Occupied / Vacant kullanılabilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                # örnek: 10 odalı gün sonu, 07.10.2026
python main.py --folyolar folyolar.xlsx --fiyatlar fiyatlar.xlsx --oda-durumu hk.xlsx \
               --tahsilat tahsilat.xlsx --kasa kasa.xlsx --tarih 07.10.2026
python main.py ... --bakiye-limiti 15000 --kasa-tolerans 1 --fiyat-tolerans 0.5 --rapor-oda-geliri 412500
```

## Çıktı

`Özet` · `Bulgular` (Yapılan İşlem sütunu boş, devir notu için) · `Kasa Mutabakatı` · `Folyolar`

## Dikkat

- **Fiyat listesi:** Liste, denetim gününe ait fiyatları içermelidir. Sezon, hafta sonu ve kampanya fiyatları ile
  sözleşmeli acente fiyatları farklıysa bunları ayrı fiyat kodu olarak verin. Listede olmayan kodlar "Bilgi" olarak
  işaretlenir.
- **Döviz:** Döviz tahsilatları için kasa ve sistem aynı para biriminde olmalıdır. Ödeme tipini döviz cinsine göre
  ayırın (örneğin "Nakit EUR").
- **Vergiler:** Konaklama vergisi ve KDV hesabı bu paketin kapsamında değildir; PMS'teki vergi ayarlarını mali
  müşavirinizle doğrulayın.
- **Kişisel veri:** Misafir adları kişisel veridir (KVKK). Çıktıyı yalnızca yetkili kişilerle paylaşın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

# Sipariş Termin Takibi · Kod Bloğu

> Satın Alma › Satın Alma Uzman Yardımcısı · Workers / Workless

ERP'den alınan açık satın alma siparişlerini termin tarihine göre izler. **Geciken** ve **gecikme riski
taşıyan** satırları tedarikçi bazında listeler. Tedarikçiye gönderilebilecek hatırlatma metni taslağı hazırlar;
hiçbir şey göndermez. İnternete bağlanmaz.

## Ne yapar?

- **Geçerli termin:** Tedarikçinin teyit ettiği termin varsa o, yoksa siparişteki termin esas alınır. Teyit
  termini istenenden sonraysa not düşülür.
- **Geciken satırlar:** Kalan miktarı olan ve geçerli termini geçmiş satırlar. Gecikme gününe göre öncelik:

  | Gecikme | Öncelik |
  |---|---|
  | 1-7 gün | Gecikmiş |
  | 8-30 gün | Yüksek |
  | 30 günden fazla | Kritik |

  Kalan tutar `kalan miktar × birim fiyat × kur` ile TL olarak hesaplanır.
- **Gecikme riski:** Şu durumlardan biri varsa satır "Riskli" olur:
  - Termin ihtiyaç (üretim/kullanım) tarihinden sonra.
  - Tahmini teslim ihtiyaç tarihinden sonra.
  - Termin önümüzdeki `--ufuk` gün (varsayılan 14) içinde ve teyit yok.
  - Termin yakın ve tedarikçi geçmişte ortalama 3 gün veya daha fazla gecikmiş.

  **Tahmini teslim** = geçerli termin + tedarikçinin geçmiş teslimlerindeki ortalama gecikme. Erken teslim 0
  sayılır; en az 3 geçmiş teslim gerekir.
- **Diğer kontroller:** Siparişten `--teyit-gun` (5) günden uzun süredir teyitsiz satırlar, kısmi ve fazla
  teslimler, termini boş veya sipariş tarihinden önce olan satırlar, kuru verilmemiş para birimleri.
- **Tedarikçi özeti:** Açık ve geciken tutar, en uzun gecikme, riskli ve teyitsiz satır sayısı, geçmiş
  zamanında teslim oranı.
- **Termin takvimi:** Önümüzdeki 6 haftada tedarikçi bazında beklenen teslim tutarı.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                     # örnek veri (rapor tarihi 08.10.2026)
python main.py --girdi acik_siparisler.xlsx --gecmis teslimler.xlsx --kur USD=41,20 EUR=48,05
python main.py --girdi acik_siparisler.xlsx --tarih 15.10.2026 --ufuk 10 --teyit-gun 3
```

| Dosya | Sütunlar |
|---|---|
| Açık siparişler | **Sipariş No, Tedarikçi, Sipariş Miktarı, Termin**; isteğe bağlı: Satır, Malzeme Kodu, Malzeme, Sipariş Tarihi, Teyit Termini, Teslim Alınan, Birim, Birim Fiyat, Para Birimi, İhtiyaç Tarihi |
| Geçmiş teslimler (`--gecmis`) | Tedarikçi, Termin, Teslim Tarihi (son 6-12 ayın teslimleri önerilir) |

- Sütun adları esnektir ("PO No", "Firma", "Gelen Miktar"...).
- Üstteki rapor başlığı satırları atlanır.
- Para birimi boşsa TL kabul edilir.

## Çıktı

`siparis_termin_takibi.xlsx`:
- `Özet`: göstergeler ve tedarikçi tablosu (geciken tutara göre sıralı).
- `Gecikenler`: öncelik ve gecikme gününe göre.
- `Gecikme Riski`.
- `Termin Takvimi`.
- `Açık Siparişler`: tüm satırlar, filtreli.
- `Tedarikçi Hatırlatma`: metin taslağı ve "Gönderildi mi?" sütunu.
- `Kontroller`.

Satır sayfalarında "Aksiyon / Yeni Termin" sütunu takip notları içindir.

## Dikkat

- **Varsayımlar:** Öncelik eşikleri ve risk kuralları örnektir; kendi satın alma prosedürünüze göre yorumlayın.
- **Gün hesabı:** Takvim günüyle yapılır; tatiller dikkate alınmaz.
- **Kurlar:** Kalan tutarlar verdiğiniz kurlarla hesaplanır. Örnek veriyle çalışırken kullanılan kurlar ve
  tarih (`ornek_veri/ayarlar.json`) gerçek değildir.
- **Hatırlatma metni:** Taslaktır. Göndermeden önce tedarikçiyle son yazışmaları kontrol edin.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçlar kontrol edilmeden işlem yapılmamalıdır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

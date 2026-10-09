# Hizmet Alımı Fatura Kontrolü · Kod Bloğu

> İdari İşler › İdari İşler Sorumlusu · Workers / Workless

Yemek, servis, temizlik ve güvenlik gibi hizmet alımı faturalarını **sözleşme birim fiyatı** ve **fiilî kullanım**
ile karşılaştırır. Kullanım kaynakları yemek kartı, servis yoklaması, puantaj ve nöbet çizelgesidir. Fazla
faturalanan tutarı hesaplar ve tedarikçiye gönderilebilecek bir **itiraz listesi** çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

### Faturalanabilir miktar

| Faturalama esası | Faturalanabilir miktar | Örnek |
|---|---|---|
| Fiilî | Kullanım toplamı. "Asgari Günlük" varsa kullanım olan her gün en az o kadar sayılır. | Öğle yemeği: asgari 120 öğün garantisi; 108 öğünlük gün 120 sayılır |
| Puantaj | Kişi-gün toplamı / "Ay Esası" (varsayılan 30) | Temizlik: 6 kişi, 1 kişi 3 gün yok → 177 / 30 = 5,9 kişi-ay |
| Sabit | Ayda 1 birim | Aylık sabit bakım bedeli |

- **Sözleşme fiyatı:** Fatura döneminin ilk günü hangi geçerlilik aralığına düşüyorsa o fiyat kullanılır. Zam
  dönemleri ayrı satır olarak girilir (ör. Ocak–Haziran 165 TL, Temmuz–Aralık 185 TL).
- **Fazla miktar:** (faturalanan − faturalanabilir) × sözleşme fiyatı. Aynı dönem ve kalem birden çok faturada varsa
  miktarlar toplanır.

### Kontroller

| Kontrol | İtiraz tutarı | Önem |
|---|---|---|
| Fazla miktar (dönem mutabakatı) | fark × sözleşme fiyatı | Yüksek |
| Birim fiyat sözleşmeden yüksek | (fatura − sözleşme fiyatı) × miktar | Yüksek |
| Sözleşmede olmayan kalem | satır tutarı | Yüksek |
| Mükerrer faturalama (aynı dönem ve kalem birden çok faturada) | mutabakatta | Yüksek |
| KDV oranı sözleşmeden farklı | — | Yüksek |
| Tevkifat tutarı = KDV × sözleşmedeki oran değil | — | Yüksek |
| Kullanım kaydı olmayan fatura (puantaj / yoklama isteyin) | — | Yüksek |
| Miktar × fiyat ≠ tutar | fazlası | Orta |
| KDV tutarı ≠ tutar × oran; tevkifat tutarı yok | — | Orta |
| Eksik faturalama; faturası gelmemiş kullanım (tahakkuk gerekebilir) | — | Bilgi |

Tutar karşılaştırmalarında 0,05 TL tolerans vardır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: Eylül 2026, 4 hizmet, 6 fatura
python main.py --sozlesme sozlesme.xlsx --kullanim kullanim.xlsx --faturalar faturalar.xlsx --donem 2026-09
```

| Dosya | Sütunlar |
|---|---|
| Sözleşme | Hizmet, Kalem, Birim, Birim Fiyat, KDV Oranı (%10 / %20), Tevkifat (ör. 5/10, 9/10), Geçerlilik Başlangıç, Geçerlilik Bitiş, Asgari Günlük, Faturalama Esası (Fiilî / Puantaj / Sabit), Ay Esası |
| Kullanım | Tarih, Hizmet, Kalem, Miktar, Kaynak. Puantaj kalemlerinde miktar kişi-gündür; ay sonuna tek satır yazılabilir. |
| Faturalar | Fatura No, Fatura Tarihi, Tedarikçi, Dönem (2026-09, 09.2026 veya Eylül 2026), Hizmet, Kalem, Miktar, Birim Fiyat, Tutar, KDV Oranı, KDV Tutarı, Tevkifat Tutarı |

Hizmet ve kalem adları üç dosyada aynı yazılmalıdır. Büyük-küçük harf ve Türkçe karakter farkı yok sayılır.

## Çıktı

`hizmet_fatura_kontrolu.xlsx`:
- `İtiraz Listesi`: tedarikçi ve fatura bazında bulgu, itiraz tutarı (KDV hariç) ve açıklama; toplam; boş
  "Tedarikçi Cevabı" sütunu.
- `Dönem Mutabakatı`: fiilî kullanım, faturalanabilir, faturalanan, fark, itiraz tutarı ve hesap açıklaması.
- `Fatura Kontrolü`: satır bazında sözleşme fiyatı, KDV, tevkifat ve beklenen tevkifat.
- `Kullanım Özeti`: kalem × ay.
- `Uyarılar`.

## Dikkat

- **Tevkifat oranları** sözleşme dosyasından okunur. KDV Genel Uygulama Tebliği'nde (I/C-2.1.3) belirlenmiş alıcılar
  için yaygın oranlar şunlardır:
  - yemek servisi 5/10;
  - temizlik, çevre ve bahçe bakımı 9/10;
  - özel güvenlik 9/10;
  - servis taşımacılığı 5/10.

  Şirketinizin tevkifat yükümlüsü olup olmadığını, tutar sınırını ve güncel oranları mali müşavirinizle kontrol edin.
- **KDV oranları** hizmete göre değişir; sözleşme dosyasına güncel oranı yazın.
- **Asgari garanti, ay esası ve eksik gün kesintisi** sözleşmenize göre girilmelidir. Paket sözleşmeyi yorumlamaz.
- **Kullanım kaynağı:** Yemek kartı, turnike, servis yoklaması ve puantaj kayıtları kesin kanıt değildir. İtirazdan
  önce kayıtların doğruluğunu teyit edin.
- **Örnek veri:** Örnek sözleşme, faturalar ve kullanım kayıtları kurgusaldır.

## Testler

Şunlar test edilir:
- Asgari günlük garanti, puantaj kişi-ay hesabı, fazla sefer ve mesai.
- Mükerrer fatura, fiyat farkı, KDV oranı, tevkifat, hesap hatası, sözleşmede olmayan kalem.
- Toplam itiraz tutarı (elle hesaplanmış).
- Fiyat geçerlilik dönemi, oran ve dönem okuma, kullanımsız fatura ve faturasız kullanım; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

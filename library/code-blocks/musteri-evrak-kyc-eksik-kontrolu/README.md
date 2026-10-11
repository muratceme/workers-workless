# Müşteri Evrak (KYC) Eksik Kontrolü · Kod Bloğu

> Şube Bankacılığı › Operasyon Yetkilisi · Workers / Workless

Müşteri dosyalarını **müşteri tipine ve kullandığı ürünlere göre** zorunlu belge listesiyle karşılaştırır; eksik, süresi
dolmuş ve süresi yaklaşan belgeleri çıkarır. Zorunlu belgeler sizin kural dosyanızdan okunur. İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Kural eşleşmesi:** Kural dosyasındaki her satır bir belgeyi bir müşteri tipi ve ürün için zorunlu kılar. `*` "tümü"
   demektir. Müşterinin tipi ve ürünlerinden biri kuralla eşleşiyorsa belge istenir. Aynı belge birden çok kuralla
   isteniyorsa en kısa geçerlilik süresi uygulanır.
2. **Geçerlilik bitişi:** Belge kaydında "Geçerlilik Bitiş" doluysa o tarih; değilse belge tarihi + kuraldaki süre (ay).
   Aynı belgeden birden çok kayıt varsa bitişi en geç olan esas alınır.
3. **Durum:**

   | Durum | Koşul |
   |---|---|
   | Eksik | Dosyada bu belge yok |
   | Süresi dolmuş | Bitiş tarihi kontrol tarihinden önce |
   | Tarih yok | Belge süreli ama tarih girilmemiş |
   | Yaklaşan | Bitişe `--yaklasan` gün (varsayılan 30) veya daha az kalmış |
   | Tamam | Diğer hâller |

4. **Sıralama:** Yüksek riskli müşteriler önce; sonra eksik, süresi dolmuş, tarihsiz ve yaklaşan belgeler.

Belge adları büyük-küçük harf ve Türkçe karakter ayırmadan eşleştirilir. Kural dosyasındaki adla dosyadaki ad aynı
olmalıdır; eşleşmeyen belgeler "Kuralda olmayan belge" uyarısıyla listelenir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --musteriler musteriler.xlsx --belgeler belgeler.xlsx --kurallar kurallar.xlsx --yaklasan 45
```

| Dosya | Sütunlar |
|---|---|
| Müşteriler | Müşteri No, Ad / Unvan, Müşteri Tipi, Ürünler (virgülle), Risk Sınıfı, Şube, Portföy Yöneticisi |
| Belgeler | Müşteri No, Belge, Belge Tarihi, Geçerlilik Bitiş |
| Kurallar | Müşteri Tipi, Ürün, Belge, Geçerlilik (ay), Açıklama |

Örnek kural satırları:

```text
Müşteri Tipi;Ürün;Belge;Geçerlilik (ay);Açıklama
Gerçek Kişi;*;Kimlik Belgesi;;
Tüzel Kişi;*;Faaliyet Belgesi;6;
*;Kredi;Gelir veya Mali Durum Belgesi;12;
```

## Çıktı

`cikti/kyc_eksik_kontrolu.xlsx`:

- `Eksik Listesi`: tamam olmayan her belge için müşteri, durum, tarihler, kalan gün, neden istendiği; boş "Talep
  Edildi" ve "Tamamlandı" sütunları.
- `Müşteri Özeti`: müşteri bazında tamamlanma oranı.
- `Belge Özeti`: en çok hangi belgenin eksik olduğu.
- `Şube Özeti`: şube ve portföy yöneticisi bazında eksikler.
- `Uyarılar`: kuralı olmayan müşteri tipi, tanımsız ürün, okunamayan tarih, listede olmayan müşteri.

## Dikkat

- **Kural dosyası sizin sorumluluğunuzdadır.** Örnek dosyadaki belgeler ve geçerlilik süreleri yalnız biçimi gösterir;
  mevzuatın veya bankanızın istediği liste değildir. Kimlik tespiti ve müşterini tanı yükümlülükleri için 5549 sayılı
  Kanun, Suç Gelirlerinin Aklanmasının ve Terörün Finansmanının Önlenmesine Dair Tedbirler Hakkında Yönetmelik, MASAK
  tebliğleri ve bankanızın uyum politikası esastır.
- **İçerik kontrolü yapılmaz.** Kod belgenin dosyada kayıtlı olup olmadığına ve tarihine bakar; belgenin doğru,
  okunaklı, imzalı veya güncel bilgi içerip içermediğini bilemez.
- **Tamam çıkması yeterlilik anlamına gelmez.** Gerçek faydalanıcının tespiti, risk değerlendirmesi ve sıkılaştırılmış
  tedbirler ayrı süreçlerdir.
- **Kişisel veri:** Müşteri listesi kişisel veri içerir. Çıktıyı yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Müşteriler, belgeler ve kurallar kurgusaldır.

## Testler

- Kural eşleşmesi (tip, ürün, `*`, en kısa süre), ay ekleme.
- Durumlar: eksik, süresi dolmuş, tarih yok, yaklaşan; bitiş sütunu; birden çok kayıt.
- Sıralama ve müşteri özeti; okuma hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

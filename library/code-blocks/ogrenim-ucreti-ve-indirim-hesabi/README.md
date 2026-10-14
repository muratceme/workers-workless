# Öğrenim Ücreti ve İndirim Hesabı · Kod Bloğu

> Kayıt Kabul ve Öğrenci İşleri › Öğrenci İşleri Sorumlusu · Workers / Workless

Özel okulun **ücret tablosu ve indirim kurallarına** göre öğrenci bazında net öğrenim ücretini, taksit planını ve kademe bazında
indirim maliyetini hesaplar. İnternete ve okul yönetim yazılımına bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Brüt ücret | Ücret tablosunda kademe + sınıf satırı; yoksa kademenin "tümü" (sınıfı boş) satırı |
| Kardeş | Kardeş sırası 2 → KARDES2; 3 ve sonrası → KARDES3 |
| Erken kayıt / personel çocuğu | Sütunda "Evet" ise ERKEN / PERSONEL |
| Peşin ödeme | Taksit sayısı 1 ise PESIN |
| Burs / diğer | Öğrenci satırındaki yüzde |
| Birleştirme | `zincir` (varsayılan): her indirim bir öncekinden kalan tutara uygulanır · `toplam`: oranlar toplanır |
| Tavan | Toplam indirim oranı `--azami` (%50) ile sınırlanır; `--burs-haric` verilirse burs tavana girmez ve ayrıca uygulanır |
| Taksit | Peşinat `--pesinat` % (taksitli ödemelerde) + kalan tutar eşit aylık taksit; son taksit kuruş farkını kapatır |

**Örnek:** 600.000 TL ücret, 2. kardeş (%10) ve erken kayıt (%5), zincirleme: 600.000 × 0,90 × 0,95 = **513.000 TL**; toplam
yöntemiyle %15 → 510.000 TL. %20 peşinat ve 8 taksitle 102.600 TL peşinat + 8 × 51.300 TL.

**Uyarılar:** okunamayan satır · ücret tablosunda olmayan kademe / sınıf (Yüksek) · tavana takılan indirim · tam burslu öğrenci (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (2026-2027, ilk taksit 01.09.2026)
python main.py --ucretler ucretler.xlsx --indirimler indirimler.xlsx --ogrenciler ogrenciler.xlsx --ilk-taksit 01.09.2026
python main.py --ucretler u.xlsx --indirimler i.xlsx --ogrenciler o.xlsx --ilk-taksit 15.08.2026 --taksit 10 --pesinat 25 --birlesim toplam --azami 40 --burs-haric
```

## Çıktı

`cikti/ogrenim_ucretleri.xlsx`: `Öğrenci Ücretleri` (tavana takılanlar sarı), `Taksit Planı`, `Kademe Özeti`, `Uyarılar`.

## Dikkat

- **Ücret ve indirimler okul politikasıdır.** Örnek tablodaki tutar ve oranlar kurgusaldır. Ücretlerin okulun ilan ettiği ücretle,
  artışların ve burs / indirim uygulamasının MEB özel öğretim kurumları mevzuatıyla uyumunu kontrol edin. Mevzuatın öngördüğü
  ücretsiz / burslu öğrenci kontenjanı ayrıca izlenmelidir.
- **Tutarlar KDV dahil kabul edilir;** vergi hesabı yapılmaz. Yemek, servis, kırtasiye gibi ek ücretler kapsam dışıdır.
- **Kişisel veri içerir.** Öğrenci ve aile bilgileri (burs, personel çocuğu) KVKK kapsamındadır.
- **Örnek veri:** Öğrenciler, ücretler ve oranlar kurgusaldır.

## Testler

- Kardeş, erken kayıt, peşin, personel, burs ve diğer indirimler; zincir ve toplam birleştirme, tavan, bursun tavan dışında tutulması.
- Peşinatlı taksit planı ve vadeler, tek taksit, kuruş kapanışı, kademe özeti; okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

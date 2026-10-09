# Masraf Fişi Kategorileme · Kod Bloğu

> Muhasebe › Muhasebe Elemanı · Workers / Workless

Çalışan masraf fişlerini ve kart harcamalarını sizin kural dosyanıza göre sınıflandırır. Gider hesabını, KDV
ayrımını ve KKEG tutarını hesaplar, muhasebe kaydı önerir. Sınıflanamayan, limit aşan ve mükerrer olabilecek
fişleri işaretler. İnternete bağlanmaz.

## Nasıl çalışır?

- **Kural dosyası:** Anahtar kelimeler (açıklamada, satıcıda veya ikisinde) → kategori, alt hesap, KDV indirilebilir
  mi, KKEG oranı, politika limiti ve not.
  - Öncelik numarası küçük olan kural önce denenir; ilk eşleşen kazanır. Örneğin "müşteri yemeği" kuralı genel
    "yemek" kuralından önce gelir.
  - Kelimeler tam kelime olarak aranır; 5 harf ve üstü kelimeler kelime içinde de bulunur.
- **Gider hesabı:** Departmanın fonksiyon hesabına alt hesap eklenir. Örneğin Satış → 760 ve yemek alt hesabı → 03
  ise `760.03` olur. Tanımsız departmanda `--varsayilan-hesap` (770) kullanılır.
- **KDV:**
  - Fatura türü belgelerde (e-Fatura, e-Arşiv, fatura, e-Bilet, serbest meslek makbuzu), kural izin veriyorsa
    191 İndirilecek KDV'ye ayrılır.
  - ÖKC (yazar kasa) fişindeki KDV gidere eklenir ve doğrulama için işaretlenir.
- **KKEG:**
  - Kuraldaki oran uygulanır; örnek kurallarda para cezaları %100, binek otomobil giderleri %30.
  - Belgesiz harcamanın tamamı KKEG olarak işaretlenir.
- **Ödeme hesabı:** Şirket Kartı → 309, Nakit (iş avansı) → 195, Kişisel Kart → 335 Personele Borçlar.
  `--odeme-hesaplari` ile değiştirilebilir.

| Kontrol | Önem |
|---|---|
| Hiçbir kurala uymayan fiş | Yüksek |
| Belgesiz harcama | Yüksek |
| Olası mükerrer: aynı çalışan, satıcı, tarih ve tutar | Yüksek |
| KDV tutarının oranla tutarsızlığı | Orta |
| Politika limiti aşımı | Orta |
| Tanımsız departman veya ödeme şekli | Orta |
| KDV'nin gidere eklenmesi, KKEG, hafta sonu harcaması | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 18 fiş, 12 kural
python main.py --masraflar masraflar.xlsx --kurallar kurallar.csv --departmanlar departmanlar.csv
python main.py --masraflar m.xlsx --kurallar k.csv --odeme-hesaplari "Şirket Kartı=300.05" "Nakit=195" --varsayilan-hesap 770
```

| Dosya | Sütunlar |
|---|---|
| Masraflar | Fiş No, Tarih, Çalışan, Departman, Satıcı, Açıklama, Tutar (KDV dahil), KDV Oranı, KDV Tutarı, Belge Türü, Ödeme |
| Kurallar | Öncelik, Anahtar Kelimeler (`\|` ile ayrılır), Alan (Açıklama / Satıcı / Hepsi), Kategori, Alt Hesap, KDV İndirilebilir (Evet / Hayır), KKEG (%), Limit (TL), Not |
| Departmanlar (isteğe bağlı) | Departman, Gider Hesabı (730 / 750 / 760 / 770…) |

Kredi kartı ekstresini veya masraf uygulamasının dışa aktarımını bu sütunlara uyarlayarak kullanabilirsiniz.

## Çıktı

`masraf_kategorileme.xlsx`:
- `Sınıflandırma`: kategori, eşleşen kelime, gider hesabı, gider, indirilecek KDV, KKEG, ödeme hesabı, bulgular; boş
  "Onay / Düzeltme" sütunu.
- `Yevmiye Önerisi`: fiş başına gider / 191 borç, ödeme hesabı alacak; toplamlar.
- `Özet`: kategori ve departman bazında tutar, gider, KDV, KKEG.
- `KKEG Listesi`: kurumlar / gelir vergisi beyannamesinde ilave edilecek kalemler için.
- `Uyarılar`.

## Dikkat

- **Kurallar ve oranlar örnektir:** KDV indirimi, KKEG oranları, binek otomobil gider kısıtlaması ve politika
  limitleri kurumunuza ve güncel mevzuata (GVK, KVK, KDVK, VUK) göre belirlenmelidir. Örnek kurallardaki mevzuat
  notları doğrulama içindir. Uygulamayı mali müşavirinizle doğrulayın.
- **Belge türü önemlidir:** ÖKC fişleri, gider pusulası ve yurt dışı faturaların vergisel sonuçları farklıdır. Yurt
  dışından hizmet alımında sorumlu sıfatıyla KDV beyanı gerekebilir.
- **Kayıt önerisi:** Yevmiye satırları öneridir; muhasebe programına aktarmadan önce kontrol edin.
- **Kişisel veri:** Masraf kayıtlarında çalışan adları bulunur. Raporu yetkili kişilerle sınırlı paylaşın.

## Testler

Şunlar test edilir:
- Kural önceliği, kelime eşleşmesi ve departman hesabı.
- Belge türüne göre KDV ayrımı; KKEG (oran, ceza, belgesiz); ödeme hesapları.
- Mükerrer, limit, KDV tutarsızlığı ve tanımsız departman uyarıları.
- Yevmiye dengesi, özel hesaplar, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

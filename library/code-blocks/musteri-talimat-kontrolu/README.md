# Müşteri Talimat Kontrolü · Kod Bloğu

> Bankacılık › Şube Bankacılığı › Gişe Yetkilisi · Workers / Workless

EFT, FAST, havale ve virman talimatlarını işleme almadan önce toplu olarak **ön kontrolden** geçirir:
IBAN, unvan, rakam-yazı tutar uyumu, işlem türü ve imza yetkisi. Hatalı ve kontrol gerektiren talimatları
işaretler. İnternete bağlanmaz.

## Ne kontrol eder?

| Alan | Kontrol | Önem |
|---|---|---|
| IBAN | Biçim, TR IBAN 26 karakter, 10. hane (rezerv) 0, mod 97 kontrol hanesi; gönderen ve alıcı için | Hata |
| Gönderen hesap | IBAN talimattaki müşteriye ait mi? Hesap para birimi talimatla aynı mı? | Hata / Dikkat |
| Gönderen unvan | Hesap unvanıyla uyuşuyor mu? "A.Ş.", "San. ve Tic." gibi ekler yok sayılır. | Dikkat |
| Alıcı unvan | Boş mu? Alıcı IBAN bankanın kendi hesap listesindeyse kayıtlı unvanla aynı mı? | Hata |
| Tutar | Rakamla tutar yazıyla tutara eşit mi? | Hata |
| İşlem türü | Havale aynı banka içinde, virman müşterinin kendi hesapları arasında mı? EFT/FAST yalnız TL mi? Aynı bankaya EFT varsa havale/virman önerilir. | Hata / Bilgi |
| Tarih | Hafta sonu verilen EFT talimatı (EFT iş günlerinde çalışır). | Bilgi |
| İmza yetkisi | İmzalayan müşterinin yetkilisi mi, yetki süresi dolmuş mu? Münferit limit, müşterek imza şartı. | Hata |
| Mükerrer | Aynı gün, aynı müşteri, aynı alıcı ve aynı tutarla birden fazla talimat. | Dikkat |

- **Yazıyla tutar:** Ayrı veya bitişik yazılabilir: "Yalnız Yüzyirmibeşbin Türk Lirası Elli Kuruş", "Bir milyon
  iki yüz bin TL", "Onbin Dolar". Okunamayan yazı "Dikkat" olarak işaretlenir.
- **İmza kuralı:**
  - Münferit yetkili kendi limitine kadar tek başına imzalayabilir.
  - Müşterek yetkilinin yanında en az bir geçerli yetkili daha olmalıdır. Müşterek imzada limit olarak
    imzalayanlardan en yüksek limitli olanın limiti esas alınır.
  - Limit boşsa sınırsız sayılır.
  - Gerçek imza sirkülerindeki grup ve limit kuralları daha ayrıntılı olabilir; yetki listesini sirküleri esas
    alarak hazırlayın.
- **Banka kodu:** Bankanızın 5 haneli kodu verilmezse en sık görülen gönderen IBAN'dan bulunur.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                          # örnek: 15 talimat
python main.py --girdi talimatlar.xlsx --hesaplar hesaplar.xlsx --yetkiler imza_yetkileri.xlsx
python main.py --girdi talimatlar.xlsx --banka-kodu 00999 --kur USD=41,20 EUR=48,05
```

| Dosya | Sütunlar |
|---|---|
| Talimatlar | Talimat No, Tarih, Tür, Müşteri No, Gönderen Unvan, Gönderen IBAN, Alıcı Unvan, Alıcı IBAN, **Tutar**, Tutar (Yazı), Para Birimi, Açıklama, İmzalayanlar (virgül, "ve" veya "/" ile ayrılmış) |
| Hesaplar (`--hesaplar`) | Müşteri No, Unvan, IBAN, Para Birimi |
| İmza yetkileri (`--yetkiler`) | Müşteri No, Yetkili, Yetki Şekli (Münferit/Müşterek), Limit (TL), Geçerlilik Bitiş |

- Hesap listesi verilmezse gönderen hesabın sahipliği ve unvan kontrolleri yapılmaz.
- Yetki listesi verilmezse imza kontrolü yapılmaz.
- Dövizli talimatlarda TL limit kontrolü için `--kur` gerekir.

## Çıktı

`talimat_kontrolu.xlsx`:
- `Özet`: hatalı, kontrol ve uygun sayıları; kontrol alanı bazında bulgular.
- `Talimatlar`: durum, bulgular ve "Kontrol Eden / Onay" sütunu.
- `Bulgular`: bulgu başına bir satır, filtreli.

## Dikkat

Bu kod bir **ön kontroldür**; aşağıdakileri yapamaz ve bankanızın prosedürünün yerini almaz:
- İmzanın imza örneğiyle karşılaştırılması.
- Kimlik ve yetki belgelerinin aslının görülmesi.
- Alıcı unvanının karşı bankadaki IBAN sahibiyle eşleşmesi. Bu kontrol bankanızın IBAN-unvan sorgu hizmetiyle
  yapılmalıdır.
- Mevzuat ve banka içi limit/saat kuralları. FAST tutar sınırı ve EFT saatleri zamanla değişir; güncel değerleri
  bankanızdan teyit edin.

Örnek verideki müşteriler, IBAN'lar ve banka kodları kurgusaldır. IBAN'lar yalnız kontrol hanesi doğru olacak
şekilde üretilmiştir.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçlar kontrol edilmeden işlem yapılmamalıdır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

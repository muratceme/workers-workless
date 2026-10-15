# Mükellef Bilgilendirme Yazısı · AI Agent

> Beyanname › Serbest Muhasebeci Mali Müşavir · Workers / Workless

Mevzuat değişikliğini ya da beyan sonucunu **mükellefin anlayacağı dille açıklayan bilgilendirme yazıları** hazırlar.
Kaynak, sizin yazdığınız konu notudur. Model notta olmayan tarih, tutar ya da madde ekleyemez. **Mükellef adı, vergi
numarası ve e-postası modele gönderilmez.**

## Nasıl çalışır?

1. **Hazırlık (kod):** Konu notu (konu satırı + numaralı bilgiler) okunur. Mükellefler mükellefiyet türüne göre
   gruplanır: kurumlar vergisi, gelir vergisi bilanço esası, serbest meslek, basit usul vb.
2. **Şablon (model):** Her grup için konunun o gruba uygulanıp uygulanmadığına karar verilir (dayandığı not
   satırlarıyla). Uygulanan gruplar için `{UNVAN}` ve `{IMZA}` yer tutuculu bir yazı şablonu yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o grubun şablonu modele geri gönderilir (varsayılan en fazla
   2 tur):
   - `{UNVAN}` ya da `{IMZA}` eksik, bilinmeyen yer tutucu var ya da yer tutucuya ek getirilmiş
   - Notta olmayan tarih, tutar ya da sayı var
   - Notta geçmeyen kanun, madde ya da tebliğ numarasına atıf yapılmış
   - Matrah, tevkifat, mahsup, tarh, tahakkuk, geçici vergi gibi bir terim ilk geçtiği cümlede açıklanmamış
   - Cümle `--cumle-azami` kelimeyi aşıyor (varsayılan 30)
   - "Ceza kesilmez", "garanti" gibi kesin vaat var
   - "Uygulanmaz" kararı bir not satırına dayanmıyor
4. **Yazılar (kod):** Şablonlar doldurulur. Yazılar gönderilmez; mükellef başına .txt ve posta birleştirme CSV'si
   yazılır. Konunun uygulanmadığı mükellefler ayrıca listelenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek geçici vergi notu
python agent.py --not konu_notu.txt --mukellefler liste.xlsx --imza "SMMM Ad Soyad · 0212 000 00 00"
```

Konu notu örneği:

```
Konu: 2026 yılı 3. dönem geçici vergi beyannamesi
1. 2026 yılı 3. dönem (Ocak–Eylül) geçici vergi beyannamesinin verilme ve ödeme süresi 17.11.2026 günü sonuna kadardır.
2. Belgelerin en geç 06.11.2026 tarihine kadar büromuza iletilmesi gerekir.
3. Basit usulde vergilendirilen mükellefler geçici vergi beyannamesi vermez.
```

## Çıktı

- `cikti/mukellef_yazilari_yazilar/`: mükellef başına yazı (.txt).
- `cikti/mukellef_yazilari.xlsx`:
  - `Şablonlar`: grup bazında karar, gerekçe, şablon, kontrol sonucu ve "Onay" sütunu.
  - `Yazılar`: mükellef bazında durum.
- `cikti/mukellef_yazilari_birlestirme.csv`: E-posta; Konu; Metin.

## Dikkat

- **Kaynak sizin notunuzdur.** Mevzuat bilgisini, tarihleri ve tutarları Resmî Gazete, Gelir İdaresi Başkanlığı ve
  TÜRMOB gibi resmî kaynaklardan teyit ederek yazın. Agent mevzuatı kendisi araştırmaz.
- **Örnek not hakkında.** Örnekteki geçici vergi bilgileri yazıldığı tarihte (Ekim 2026) kaynaklardan teyit edildi.
  Dördüncü geçici vergi dönemi 7338 sayılı Kanun'la 2022'den itibaren kaldırılmış, 7566 sayılı Kanun'la (RG 19.12.2025)
  yeniden getirilmiştir. Süreler mücbir sebep ya da idari kararlarla uzatılabilir; göndermeden önce güncel durumu kontrol
  edin.
- **Kişisel veri ve sır saklama.** Mükellef bilgileri modele gönderilmez. Çıktı dosyalarını yalnız yetkili büro
  personeli kullanmalıdır.
- **Örnek veri:** Mükellefler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek:
- kurumlar vergisi şablonunda imzayı unutur ve yer tutucuya ek getirir,
- serbest meslek şablonuna notta olmayan "VUK 359. madde" ve "1.500 TL ceza" ekleyip "ceza kesilmez" der,
- basit usul için "uygulanmaz" kararını not satırına dayandırmaz.

Kodun bunları yakaladığı, yalnız sorunlu grupları yeniden yazdırdığı, açıklanmamış terimi ve uzun cümleyi tespit ettiği
doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Vergi danışmanlığının yerine geçmez. API kullanım ücretleri ve gönderilen yazıların
doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

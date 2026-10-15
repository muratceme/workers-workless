# Aday Bilgilendirme E-postaları · AI Agent

> İnsan Kaynakları › İnsan Kaynakları Uzman Yardımcısı · Workers / Workless

Aday listesinden **mülakat daveti, olumsuz dönüş ve iş teklifi e-postaları** hazırlar. **Aday adı, e-posta adresi, ücret
ve tarihler yapay zekâ modeline gönderilmez.** Model yalnız aşama türleri ve pozisyon adları için yer tutuculu şablon
yazar; bilgileri kod doldurur. E-posta göndermez; posta birleştirme için CSV üretir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Her satırın aşaması tanınır ve kombinasyonlar çıkarılır: yüz yüze / online davet, başvuru
   aşamasında / mülakat sonrasında olumsuz dönüş, teklif. Satır kontrolleri:
   - E-posta adresi geçersiz
   - Davette tarih, saat ya da yer; teklifte ücret, başlangıç ya da yanıt son tarihi eksik
   - Mülakat tarihi geçmiş

   Bu satırlar "Gönderilemeyen" sayfasına düşer.
2. **Şablon (model):** Her kombinasyon için konu satırı ve metin yazılır: `{AD}`, `{POZISYON}`, `{TARIH}`, `{SAAT}`,
   `{YER}`, `{UCRET}`, `{BASLANGIC}`, `{SON_YANIT}`, `{SIRKET}`, `{IMZA}`, `{ILETISIM}`.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o şablon modele geri gönderilir (varsayılan en fazla 2 tur):
   - Zorunlu yer tutucu eksik ya da bilinmeyen yer tutucu var
   - Yer tutucuya ek getirilmiş ("{AD}'a")
   - Şablonda sabit sayı, tarih ya da saat var
   - Ayrımcılık sayılabilecek ifade var: yaş, cinsiyet, medeni durum, hamilelik, din, etnik köken, engellilik, sağlık,
     uyruk, askerlik
   - Olumsuz dönüşte tutulamayacak söz var ("ileride mutlaka sizinle çalışacağız")
   - Konu satırı `--konu-azami` karakteri aşıyor (varsayılan 80)
4. **Doldurma (kod):** Kontrolden geçen şablonlarla e-postalar oluşturulur.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek aday listesi
python agent.py --girdi adaylar.xlsx --firma firma.txt
python agent.py --girdi adaylar.xlsx --firma firma.txt --konu-azami 60 --bugun 15.10.2026
```

Firma dosyasında `Şirket: …`, `İmza: …` ve `İletişim: …` satırları bulunmalıdır. Üslup ve süreç notlarınızı da
yazabilirsiniz (ör. "Olumsuz dönüşlerde gerekçe yazmıyoruz").

## Çıktı

- `cikti/aday_epostalari.xlsx`:
  - `Şablonlar`: konu ve metin, kontrol sonucu, "Onay" sütunu.
  - `E-postalar`: aday başına doldurulmuş konu ve metin.
  - `Gönderilemeyen`: eksik bilgi, geçersiz e-posta, geçmiş tarih ya da kontrolden geçmeyen şablon.
- `cikti/aday_epostalari_birlestirme.csv`: E-posta; Konu; Metin (Outlook / Gmail posta birleştirme için).

## Dikkat

- **Gönderim sizde.** Şablonları İK sorumlusu onaylamalıdır. Teklif e-postası yazılı teklif ya da iş sözleşmesinin
  yerine geçmez.
- **Ayrımcılık yasağı.** İş ilişkisinde ayrımcılık 4857 sayılı İş Kanunu md. 5 ve 6701 sayılı Türkiye İnsan Hakları ve
  Eşitlik Kurumu Kanunu ile yasaklanmıştır. Kod belirgin ifadeleri yakalar; son kontrol sizdedir.
- **Kişisel veri.** Aday bilgileri modele gönderilmez. Adaylara KVKK aydınlatma metninizi başvuru sırasında
  sunmuş olmalısınız; CSV ve Excel çıktısını yalnız yetkili personel kullanmalıdır.
- **Örnek veri:** Adaylar ve şirket kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek:
- daveti "45 dakika" diye sabit sayıyla yazar,
- olumsuz dönüşte yaşa değinir,
- yer tutucuya ek getirip "ileride mutlaka sizinle çalışacağız" der,
- teklifte yanıt son tarihini unutur.

Kodun bunları yakaladığı, yalnız sorunlu şablonları yeniden yazdırdığı ve aday bilgisini modele göndermediği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve gönderilen e-postaların doğruluğu kullanıcının sorumluluğundadır.
Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

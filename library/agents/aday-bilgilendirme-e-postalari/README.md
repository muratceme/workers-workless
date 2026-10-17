# Aday Bilgilendirme E-postaları · AI Agent

> İnsan Kaynakları › İnsan Kaynakları Uzman Yardımcısı · Gerisi Bende

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

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |
| 6701 sayılı Türkiye İnsan Hakları ve Eşitlik Kurumu Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6701.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.

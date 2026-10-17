# KVKK Aydınlatma Metni Taslağı · AI Agent

> Hukuk › Avukat · Gerisi Bende

Veri işleme envanterinden **6698 sayılı KVKK md. 10'a uygun aydınlatma metni taslağı** hazırlar. Amaç, hukuki sebep,
aktarım ve toplama yöntemi bölümlerini model yazar. **Veri sorumlusunun kimliği, md. 11 hakları ve başvuru yöntemleri
kod tarafından aynen eklenir.** Envanterde olmayan bir hukuki sebep ya da "vb." gibi genel ifade yazılırsa taslak
modele geri gönderilir.

## Nasıl çalışır?

1. **Envanter kontrolü (kod):** Her satırın hukuki sebebi KVKK bendine çevrilir ("5/2-c", "md. 5/2 (ç)", "Açık rıza" →
   özel nitelikli veride 6/3-a). Aşağıdakiler "Envanter Kontrol" sayfasında işaretlenir:
   - Hukuki sebep tanınmıyor (md. 5/1, 5/2-a … f, 6/3-a … g dışında)
   - Özel nitelikli veri (sağlık, biyometrik, din, ceza mahkûmiyeti, sendika üyeliği …) md. 5'e dayandırılmış
     (7499 sayılı Kanun'la değişik md. 6/3 şartlarından biri gerekir)
   - Yurt dışı aktarımda md. 9 mekanizması yok (yeterlilik kararı, standart sözleşme, bağlayıcı şirket kuralları,
     taahhütname, arızi aktarım)
2. **Taslak (model):** Giriş, amaç bölümleri (her biri dayandığı E satırlarıyla), aktarım ve toplama yöntemi.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bir envanter satırı hiçbir bölümde yok; bölüm, dayandığı satırın amacını ya da hukuki sebebini yazmamış
   - Metinde envanterde olmayan bir hukuki sebep geçiyor (ör. meşru menfaate dayanan veri için "açık rızanıza dayanılarak")
   - Bir alıcı grubu, yurt dışı aktarım ve mekanizması ya da bir toplama yöntemi yazılmamış
   - "vb.", "ve benzeri", "sınırlı olmamak üzere" gibi genel ifadeler; "okudum, onaylıyorum" gibi açık rıza beyanı
4. **Metin (kod):** Veri sorumlusu (unvan, MERSİS, adres), md. 11'deki haklar ve Veri Sorumlusuna Başvuru Usul ve
   Esasları Hakkında Tebliğ'deki başvuru yolları (yazılı, KEP, güvenli elektronik imza / mobil imza, kayıtlı e-posta;
   en geç 30 gün, kural olarak ücretsiz) eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                # örnek (kurgusal) e-ticaret sitesi
python agent.py --envanter envanter.xlsx --sorumlu sirket.txt
```

Envanter örneği (`;` ile ayrılmış):

| Veri Kategorisi | Veriler | İşleme Amacı | Hukuki Sebep | Toplama Yöntemi | Alıcı Grubu | Aktarım Amacı | Yurt Dışı Aktarım |
|---|---|---|---|---|---|---|---|
| Kimlik | Ad, soyad | Üyelik ve sipariş süreçleri | 5/2-c | Üyelik formu | Kargo şirketleri | Teslimat | |
| İşlem Güvenliği | IP adresi | Bilgi güvenliği | 5/2-f | Otomatik | Bulut barındırma hizmeti sağlayıcısı | Sunucu hizmeti | Standart sözleşme |

## Çıktı

- `cikti/aydinlatma_metni.txt`: yayına hazırlanacak metin taslağı.
- `cikti/aydinlatma_metni.xlsx`: Metin (bölüm ve dayanaklar), Envanter Kontrol.

## Dikkat

- **Hukuki değerlendirme sizde.** Hukuki sebebin doğru seçilmesi (ör. meşru menfaat için denge testi) envanterin
  sorumluluğundadır; agent yalnız envantere sadık metin yazar. Yayından önce avukat / KVKK sorumlusu incelemelidir.
- **Açık rıza ayrı alınır.** Aydınlatma metni rıza beyanı içermez; açık rıza gereken işlemler için ayrı metin hazırlayın.
- **Kapsam:** Kurum karar ve rehberleri ile sektörel mevzuat (ör. sağlık, bankacılık, ticari elektronik ileti)
  ayrıca dikkate alınmalıdır. Kanunun 7499 sayılı Kanun'la (01.06.2024) değişen md. 6 ve 9 hükümleri esas alınmıştır;
  mevzuatın güncel hâlini kontrol edin.
- **Örnek veri:** Şirket kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta meşru menfaate dayanan güvenlik kayıtlarını "açık rızanıza
dayanılarak" diye yazar, "vb." kullanır, hukuki süreç satırını atlar, yurt dışı aktarımı yazmaz ve sona "okudum,
onaylıyorum" ekler. Kodun beş sorunu da yakaladığı, envanter kontrolünün özel nitelikli veriyi ve tanınmayan hukuki
sebebi işaretlediği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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

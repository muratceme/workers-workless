# Hasar İhbarı Ön Kayıt · AI Agent

> Hasar Yönetimi › Hasar Uzman Yardımcısı · Workers / Workless

E-posta, web formu ya da telefon notuyla gelen hasar ihbarlarından poliçe no, hasar tarihi, yeri, türü, tahmini tutar,
plakalar ve tutanak bilgisini ayıklayıp **ön kayıt** tablosuna çevirir. Çıkarılan bilgiler kodla ihbar metnine ve poliçe
listesine karşı denetlenir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Metindeki plakalar Türkiye plaka biçimine göre bulunur. Telefon, e-posta, TCKN ve IBAN maskelenir.
2. **Ön kayıt (model):** Her ihbar için alanlar doldurulur ve dayanılan ifade metinden alıntılanır. Eksik bilgiler ve
   sigortalıdan istenecek belgeler listelenir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa ihbar modele geri gönderilir (varsayılan en fazla 2 tur):
   - Poliçe no ya da tutar metinde geçmiyor
   - Hasar tarihi metinde yazmıyor ve "dün / bugün / önceki gün" ifadesiyle ihbar tarihinden de çıkarılamıyor
   - Hasar tarihi ihbar tarihinden sonra
   - Plaka metindeki plakalardan değil ya da sigortalı ve karşı taraf plakası aynı
   - Metinde yaralanma / hastane ifadesi var ama yaralanma işaretlenmemiş
   - Alıntı metinde yok
4. **Poliçe ve süre kontrolü (kod):**
   - Poliçe listede yok
   - Hasar tarihi poliçe süresi dışında
   - Branş ya da plaka poliçeyle uyuşmuyor
   - Poliçe no yazılmamışsa plakadan poliçe önerisi
   - İhbar, hasardan `--ihbar-gun` iş gününden geç yapılmış (varsayılan 5)
   - Yaralanma bildirilen ihbarlar "Öncelikli" olarak en üste alınır

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek ihbarlar + poliçe listesi
python agent.py --girdi ihbarlar.xlsx --policeler policeler.xlsx
python agent.py --girdi ihbarlar.xlsx --ihbar-gun 5 --tur 3
```

## Çıktı

- `cikti/hasar_on_kayit.xlsx`:
  - `Ön Kayıt`: alanlar, öncelik, eksik bilgiler, istenecek belgeler, kontrol uyarıları ve "Onay" sütunu.
  - `İhbar Metinleri`: orijinal metin ve bulunan plakalar.

## Dikkat

- **Ön kayıttır.** Hasar dosyası açılmadan önce ihbar ve poliçe bilgileriyle karşılaştırılmalıdır. Teminat, kusur ve ödeme
  değerlendirmesi yapmaz.
- **İhbar süresi.** Kasko genel şartlarında rizikonun öğrenilmesinden itibaren 5 iş günü içinde bildirim öngörülür; diğer
  branşlarda süre poliçe şartlarına göre değişir. `--ihbar-gun` değerini ürününüze göre ayarlayın. Hesap hafta sonlarını
  dışarıda bırakır, resmî tatilleri hesaba katmaz. Sürenin aşılması tek başına ret nedeni değildir; değerlendirmeyi hasar
  uzmanı yapar.
- **İstenecek belgeler** genel bir listedir; şirketinizin hasar belge listesine göre kontrol edin.
- **Veri gönderimi.** İhbar metinleri seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon, e-posta, TCKN ve IBAN
  maskelenir; ad, adres ve plakalar maskelenmez. Gerekirse yerel model (Ollama) kullanın.
- **Örnek veri:** Kişiler, poliçeler ve plakalar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek "dün" yazan ihbarda yanlış tarih verir, hastanede olan
sigortalının yaralanmasını atlar ve metinde olmayan bir poliçe numarası yazar. Kodun bunları yakalayıp düzelttirdiği, poliçe
süresi dışındaki hasarı ve geç ihbarı bulduğu, plakadan poliçe önerdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

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

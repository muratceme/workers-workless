# Dilekçe Taslağı · AI Agent

> Hukuk Bürosu › Dava Takip › Avukat · Workers / Workless

Avukatın olay özeti, talepler ve deliller listesinden dava dilekçesi taslağı hazırlar. Dilekçenin unsurlarını HMK
md. 119'a göre kontrol eder ve modelin yazdığı her vakıayı olay özetine dayandırır. **Taslaktır: avukat kontrolü ve
imzası olmadan kullanılamaz.** Kendi API anahtarınızla çalışır.

## Nasıl çalışır?

1. **Girdiler:** Dava bilgileri, olay özeti, talepler ve deliller okunur. `mevzuat/` klasörüne kanun maddesi
   metinleri koyarsanız maddelere bölünür.
2. **Kod kontrolü:** Kod, HMK md. 119/1'de sayılan unsurlara bakar:

| Bent | Unsur |
|---|---|
| a | mahkemenin adı |
| b | tarafların adı ve adresi |
| c | davacının T.C. kimlik numarası (gerçek kişi) veya vergi no / MERSİS (tüzel kişi); TCKN ve VKN kontrol haneleri doğrulanır |
| ç | vekilin adı ve adresi |
| d | dava konusu ve dava değeri |
| e | vakıaların özeti |
| f | deliller |
| g | hukuki sebepler |
| ğ | talep sonucu |
| h | imza |

   - **Arabuluculuk:** Dava türü zorunlu arabuluculuğa tabi olabilecek bir türdeyse (ticari alacak, iş, tüketici,
     kira vb.), arabuluculuk son tutanağının dosyada olup olmadığına bakar ve uyarır.
3. **Maskeleme:**
   - Taraf ve vekil adları `[DAVACI]`, `[DAVALI]`, `[DAVACI VEKİLİ]` olarak; `--gizle` ile verilen adlar `[KİŞİ-n]`
     olarak gönderilir.
   - Adresler ve kimlik numaraları modele hiç gönderilmez.
   - Telefon, e-posta, TCKN ve IBAN maskelenir.
   - Göndermeden önce onay istenir.
4. **Model yazar:**
   - açıklamalar: her vakıa için olay özetinden birebir dayanak alıntısı ve delil numaraları
   - hukuki sebepler
   - sonuç ve istem
   - avukata notlar: eksik bilgi, olası savunmalar, usul kontrolleri
5. **Kod denetler:**
   - **Vakıalar:** Dayanak alıntısı olay özetinde birebir geçmiyorsa vakıa "dayanak doğrulanamadı" diye işaretlenir.
   - **Deliller:**
     - Listede olmayan delil numaraları silinir.
     - Delilsiz vakıa (HMK md. 119/1-f) ve kullanılmayan delil uyarılır.
   - **Hukuki sebepler:**
     - Verilen mevzuatta olmayan madde numarasına atıf uyarılır.
     - Kanunlar listesinde olmayan kanun sayısı uyarılır.
   - **Sayılar:** Girdilerde olmayan tarih veya tutar uyarılır.
   - **Sonuç ve istem:** Dava değerinin sonuç ve istemde geçmemesi ve eksik talep uyarılır.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                          # örnek: kurgusal ticari alacak davası
python agent.py --girdi ./dava_klasoru --gizle "Tanık Adı" "Şirket Yetkilisi"
```

| Dosya | İçerik |
|---|---|
| `dava_bilgileri.txt` | "Alan: değer" satırları: Mahkeme, Davacı, Davacı TCKN veya Davacı Vergi No / MERSİS No, Davacı Adres, Davacı Vekili, Vekil Adres, Davalı, Davalı Vergi No, Davalı Adres, Dava Türü, Dava Değeri, Konu |
| `olay_ozeti.txt` (.docx / .pdf) | Avukatın kronolojik olay anlatımı: tarihler, tutarlar, belgeler |
| `talepler.txt` | Talepler, satır satır |
| `deliller.csv` | Delil No, Delil, Tür, Ek |
| `mevzuat/` (isteğe bağlı) | Dayanılacak kanun maddelerinin metni ("Madde 117 – …" biçiminde) |

## Çıktı

`dilekce_taslagi.md`:
- **Dilekçe:** başlık, taraflar, dava değeri, konu, numaralı açıklamalar (delil atıflarıyla), hukuki sebepler,
  hukuki deliller, sonuç ve istem, imza yeri ve ekler.
- **Dilekçeye dahil olmayan bölüm:** avukata notlar ve kod kontrolleri, dilekçenin altında ayrı yer alır.

`dilekce_taslagi.xlsx`:
- **Sayfalar:**
  - `Kontrol Listesi`: HMK 119 unsurları
  - `Vakıa–Delil`: vakıa, dayanak alıntısı, doğrulama, deliller
  - `Hukuki Sebepler`
  - `Kontroller`
  - `Avukata Notlar`
- **Hücre renkleri:** Model metinleri mor hücrededir; avukat kontrol sütunları sarıdır.

## Dikkat

- **Avukat kontrolü zorunludur:** Bu bir taslaktır; hukuki görüş değildir. Avukat şunlara kendisi karar verir:
  - görevli ve yetkili mahkeme
  - dava şartları (arabuluculuk vb.)
  - zamanaşımı ve hak düşürücü süreler
  - harç ve gider avansı (araç hesaplamaz)
  - faiz türü
  - ihtiyati tedbir / haciz talepleri
- **Mevzuat ve içtihat:** Model yalnız `mevzuat/` klasöründe verilen maddelere numarayla atıf yapabilir. Yargıtay
  kararı veya içtihat yazması yasaktır. Madde metinlerini güncel resmî kaynaktan (mevzuat.gov.tr) alın.
- **Arabuluculuk uyarısı yalnız bir ipucudur:** Hangi uyuşmazlıkların dava şartı arabuluculuğa tabi olduğu kanunlarla
  belirlenir ve değişebilir. Dava türü için güncel mevzuatı kontrol edin.
- **Veri gizliliği:** Müvekkil bilgileri avukatlık sır saklama yükümlülüğü ve KVKK kapsamındadır. Gönderim öncesi
  maskelemeyi kontrol edin; gerekirse `--gizle` ile ek adları maskeleyin veya yerel model (Ollama) kullanın.
- **Örnek veri:** Örnek dava, taraflar ve tutarlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz; model yanıtı sahte bir fonksiyonla verilir. Şunlar test edilir:
- Okuma ve HMK 119 kontrol listesi; TCKN / VKN doğrulama; arabuluculuk uyarısı.
- Maskeleme: ad, adres ve kimlik numarası gönderilmez.
- Dayanak alıntısı doğrulama; delil numaraları; madde ve kanun atıfları; girdide olmayan sayılar.
- Dilekçe biçimi, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6100 sayılı Hukuk Muhakemeleri Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6100.pdf) | Resmî metinle karşılaştırılan maddeler: 127, 136, 281, 345, 347, 361 |
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

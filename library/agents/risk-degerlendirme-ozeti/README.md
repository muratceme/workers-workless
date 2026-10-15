# Risk Değerlendirme Özeti · AI Agent

> Sigorta › Teknik (Risk Kabul / Underwriting) › Teknik Uzman (Underwriter) · Workers / Workless

Risk inceleme (survey) raporundan **risk değerlendirme özeti** ve **risk kabul önerisi taslağı** hazırlar.
Yangın, deprem, sel/su baskını, hırsızlık ve sorumluluk risklerini, rapordaki iyileştirme önerilerini ve
öncelikleri derler.
- **Kod:** kontrol listesini, olumsuz gözlemleri ve bedelleri kontrol eder.
- **Model:** her gözlemi rapordan **birebir alıntıyla** yazar; kod bu alıntıları doğrular.

Kabul kararı yetkili underwriter'a aittir.

> **Önemli — gizlilik:** Survey raporu sigortalıya ait ticari bilgiler içerir. Dış bir yapay zekâ servisine
> göndermeden önce şirketinizin onayını alın. Veri bilgisayardan çıkmasın isteniyorsa `.env` içinde
> `WW_PROVIDER=ollama` ile yerel model kullanın.

## Nasıl çalışır?

1. **Okur (kod):**
   - Survey raporu ve ekleri: `.txt`, `.pdf`, `.docx`.
   - İsteğe bağlı sigorta bedelleri tablosu: `Kalem`, `Bedel`.
2. **Kontrol eder (kod):**

   | Kontrol | Önem |
   |---|---|
   | Kontrol listesindeki yangın veya deprem konusu raporda geçmiyor | orta |
   | Kontrol listesindeki diğer konular raporda geçmiyor | bilgi |
   | Koruma önlemi "yok", "bulunmamaktadır", "eksik", "yapılmamıştır", "ankrajlı değildir" gibi geçiyor (olumsuz gözlem) | orta |
   | Emtia bedeli raporda yazan maksimum stok değerinin altında (eksik sigorta riski) | orta |
   | Bedeli yazılmamış kalem | orta |

   **Kontrol listesi** 25 konudan oluşur:
   - Yangın: yapı tarzı, çatı/cephe, algılama, sprinkler, yangın dolabı/hidrant, yangın suyu, elektrik/termal
     ölçüm, yanıcı maddeler, sıcak çalışma, yangın duvarı, itfaiye, yıldırımdan korunma, düzen-temizlik.
   - Deprem: bina yaşı, zemin etüdü, raf ankrajı, ruhsat/iskân.
   - Sel: dereye mesafe, bodrum/drenaj.
   - Hırsızlık: çevre güvenliği, güvenlik görevlisi, kamera/alarm.
   - Sorumluluk: çalışan/İSG, ziyaretçi.
   - Diğer: hasar geçmişi.

   **Olumsuz gözlem taraması:** Yalnız koruma önlemleri (sprinkler, dedektör, alarm, yangın duvarı, ankraj gibi)
   için yapılır. "Bodrum kat yoktur" gibi bir tehlikenin yokluğu olumsuz sayılmaz.
3. **Maskeler:**
   - Sigortalı unvanı `[SİGORTALI]`, risk adresi `[ADRES]` olur.
   - Görüşülen kişi ve `--gizle` ile verilen adlar `[KİŞİ-n]` olur.
   - Telefon, e-posta, TCKN ve IBAN maskelenir.
   - Veri gönderilmeden önce onayınız alınır.
4. **Yazar (model):**
   - Tesis özeti.
   - Her risk için seviye (Yüksek / Orta / Düşük / Bilgi yok), olumlu ve olumsuz gözlemler (alıntılı) ve
     değerlendirme.
   - Tüm rapor önerileri, öncelikleriyle (Kabul öncesi / Kısa / Orta / Uzun vade).
   - Kabul önerisi ve şartlar; eksik bilgiler.
5. **Denetler (kod):**
   - Alıntı raporda birebir geçmiyor → "doğrulanamadı".
   - Rapordaki numaralı bir öneri (`R1.`, `Öneri 1.`) özette yok → uyarı.
   - Yüksek risk veya kabul öncesi iyileştirme varken öneri "Kabul" → uyarı.
   - "Şartlı kabul" önerisinde şart yok → uyarı.
   - Metinde, girdilerde olmayan sayı → işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                          # örnek: kurgusal plastik ambalaj fabrikası
python agent.py --girdi ./risk_klasoru
python agent.py --girdi ./risk_klasoru --gizle "Görüşülen Kişi"
```

**Rapor biçimi:** Raporda `Sigortalı:`, `Risk Adresi:` ve `Görüşülen:` satırları varsa maskeleme otomatik yapılır.
Öneriler `R1.`, `R2.` gibi numaralıysa kapsama kontrolü yapılır.

## Çıktı

- **Özet:** `risk_degerlendirme_ozeti.md`, takma adlar geri açılmış.
- **Excel:** `risk_degerlendirme_ozeti.xlsx`:
  - `Özet`: boş "Underwriter Kararı" ve "Karar Notu" satırları.
  - `Riskler`: gözlem, alıntı, doğrulandı mı.
  - `İyileştirmeler`: öncelik sırasıyla; "Takip / Sigortalı Cevabı" sütunu.
  - `Kontrol Listesi`: konu raporda var mı, ilk geçtiği cümle.
  - `Kontroller`.

## Dikkat

- **Taslak niteliği:** Risk seviyeleri ve kabul önerisi taslaktır. Tarife, teminat, muafiyet ve özel şartları
  şirketinizin risk kabul kurallarına ve güncel tarifelere göre belirleyin.
- **Kontrol listesi:** Kelime aramasıyla çalışır. Konunun raporda "geçmesi" yeterli ve doğru bilgi olduğunu
  göstermez; "Raporda bilgi yok" uyarısı da farklı ifade edilmiş bir bilgiyi kaçırmış olabilir.
- **Eksik sigorta uyarısı:** Raporda "maksimum stok … TL" gibi açık bir değer varsa çalışır. Değer esaslarını
  (yeni değer, rayiç değer vb.) poliçe şartlarına göre ayrıca değerlendirin.
- **Örnek veri:** Örnek tesis, kişiler ve rapor kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Bedeller ve eksik sigorta uyarısı; kontrol listesi (yıldırımdan korunma eksik).
- 9 olumsuz gözlem; "Bodrum kat yoktur" ve öneri satırlarının sayılmaması.
- Maskeleme ve kişi/firma/adresin modele gitmemesi.
- Alıntı doğrulama; özette olmayan R7.
- "Kabul" ile yüksek risk çelişkisi; uydurma sayı (250.000 TL).

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

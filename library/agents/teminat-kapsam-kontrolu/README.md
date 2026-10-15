# Teminat Kapsam Kontrolü · AI Agent

> Sigorta › Hasar Yönetimi › Hasar Uzmanı · Workers / Workless

Hasar olayını poliçe teminatları, istisnalar, muafiyetler ve özel şartlarla karşılaştırıp **kapsam değerlendirmesi
taslağı** hazırlar.
- **Kod:** süre, prim ve teminat kontrollerini yapar; eksik sigorta, muafiyet ve limitle tazminat ön hesabını çıkarır.
- **Model:** her değerlendirmeyi poliçeden ve olay belgelerinden birebir alıntıyla destekler; kod alıntıları doğrular.

Kapsam ve tazminat kararı yetkili hasar uzmanına aittir.

> **Önemli — gizlilik:** Hasar dosyası kişisel veri içerir. Dış bir yapay zekâ servisine göndermeden önce şirketinizin
> onayını alın. Veri bilgisayardan çıkmasın isteniyorsa `.env` içinde `WW_PROVIDER=ollama` ile yerel model kullanın.

## Nasıl çalışır?

1. **Okur (kod):**

   | Dosya | Nasıl tanınır |
   |---|---|
   | Hasar bilgileri | `hasar_bilgileri.txt` |
   | Teminat tablosu | `Teminat`, `Limit`, `Muafiyet` sütunları |
   | Poliçe / şart metinleri | Adında "poliçe", "şart", "kloz" geçen dosyalar |
   | Olay belgeleri | İhbar, eksper raporu, tutanak, beyan |

2. **Kontrol eder (kod):**

   | Kontrol | Önem |
   |---|---|
   | Hasar tarihi poliçe süresi dışında | yüksek |
   | Prim ödenmemiş | yüksek |
   | Talep edilen teminat poliçede yok | yüksek |
   | Hasar, poliçe başlangıcından sonraki 30 gün içinde | orta |
   | Eksik sigorta: gerçek değer sigorta bedelinden yüksek | orta |
   | Poliçe metni verilmedi | orta |
   | İhbar hasardan kaç gün sonra yapılmış | bilgi |

3. **Ön hesap yapar (kod):** hasar tutarı → kapsam dışı kalemler → eksik sigorta oranı (bedel / gerçek değer) →
   muafiyet → limit.
   - Muafiyet yazımı: `%2, en az 5.000 TL`, `%10`, `10.000 TL`, `%5 en az 1.000 en fazla 20.000`.
4. **Maskeler:**
   - Sigortalı `[SİGORTALI]`, poliçe no `[POLİÇE NO]` olur.
   - Metinlerde "yetkilisi", "Görüşülen", "Beyan veren", "Sürücü" ifadelerinden sonra gelen adlar ve `--gizle` ile
     verilen adlar `[KİŞİ-n]` olur.
   - Telefon, e-posta, TCKN ve IBAN maskelenir.
   - Veri gönderilmeden önce onayınız alınır.
5. **Yazar (model):**
   - Olay özeti.
   - Teminat değerlendirmesi: Kapsamda / Kapsam dışı / Belirsiz.
   - İstisnalar ve şartlar: Uygulanır / Uygulanmaz / Belirsiz.
   - Kapsam dışı kalabilecek kalemler, tutarlarıyla.
   - Sonuç: Kapsamda / Kısmen kapsamda / Kapsam dışı / Ek bilgi gerekli.
   - Belgeler arası çelişkiler ve ek bilgi ihtiyacı.
6. **Denetler (kod):**

   | Durum | Sonuç |
   |---|---|
   | Poliçe veya olay alıntısı belgede birebir geçmiyor | "doğrulanamadı" |
   | Kapsam dışı kalemin tutarı veya cümlesi belgelerde yok | ön hesaba alınmaz |
   | Kapsam dışı kalemler doğrulandıysa | ikinci ön hesap yapılır |
   | "Kapsamda" sonucu yüksek bir kod bulgusuyla veya açık bir istisnayla çelişiyor | uyarı |
   | Metinde, girdilerde olmayan sayı | işaretlenir |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                    # örnek: kurgusal işyeri dahili su hasarı
python agent.py --girdi ./hasar_klasoru
python agent.py --girdi ./hasar_klasoru --gizle "Eksper Adı"
```

`hasar_bilgileri.txt` satırları: `Dosya No:`, `Sigortalı:`, `Poliçe No:`, `Poliçe Başlangıç:`, `Poliçe Bitiş:`,
`Hasar Tarihi:`, `İhbar Tarihi:`, `Prim Durumu:`, `Talep Edilen Teminat:`, `Hasar Tutarı:`, `Sigorta Bedeli:`,
`Gerçek Değer:`.

## Çıktı

- **Değerlendirme:** `teminat_kapsam.md`, takma adlar geri açılmış.
- **Excel:** `teminat_kapsam.xlsx`:
  - `Özet`: boş "Hasar Uzmanı Kararı" ve "Karar Notu" satırları.
  - `Teminat ve İstisnalar`: sonuç, gerekçe, alıntılar; "İnceleme" sütunu.
  - `Ön Hesap`: iki senaryo, adım adım.
  - `Kontroller`.

## Dikkat

- **Taslak niteliği:** Değerlendirme taslaktır. Poliçenin genel ve özel şartlarının tamamını, klozları ve güncel
  mevzuatı esas alın; model yalnız verdiğiniz metinleri görür.
- **Ön hesap:** Basitleştirilmiştir. Değer esası (yeni değer / rayiç değer), muafiyetin hangi tutara
  uygulanacağı, işlem sırası ve sovtaj poliçeye göre değişebilir.
- **İhbar ve prim:** İhbar süresi ve primin ödenmemesinin sonuçları için ilgili genel şartlara bakın; kod yalnız
  gün sayısını ve durumu gösterir.
- **Örnek veri:** Örnek poliçe, şartlar, kişiler ve olay kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Belge tanıma; muafiyet ayrıştırma; ön hesap (85.000 TL) ve ikinci senaryo (32.500 TL).
- Süre dışı hasar, ödenmemiş prim, poliçede olmayan teminat.
- Maskeleme.
- Doğrulanamayan alıntı ve kalem; "Kapsamda" çelişkisi; uydurma sayı.

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

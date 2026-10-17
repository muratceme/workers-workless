# Firma İstihbarat Raporu Derleme · AI Agent

> Bankacılık › Şube Bankacılığı › Ticari Portföy Yöneticisi · Gerisi Bende

Ticaret sicili, banka istihbaratı, piyasa görüşmeleri, KKB/memzuç riski ve olumsuz kayıtları **tek bir
istihbarat raporunda** toplar.
- **Kod:** Tutarlılık kontrollerini yapar ve tabloları hazırlar.
- **Model:** Kaynaklara dayanarak derlemeyi yazar; kaynaklar arası çelişkileri ve teyit edilecek noktaları
  listeler.

Kredi kararı vermez.

> **Önemli — bankacılık sırrı:** İstihbarat bilgileri bankacılık sırrı ve ticari sır niteliğindedir. Dış bir
> yapay zekâ servisine göndermek bankanızın politikasına aykırı olabilir; bilgi güvenliği ve uyum birimlerinizin
> onayını alın. Veri bilgisayardan çıkmasın isteniyorsa `.env` içinde `WW_PROVIDER=ollama` ile yerel model
> kullanın.

## Nasıl çalışır?

1. **Dosyaları tanır (kod):**
   - Ticaret sicili özeti: adında "sicil" geçen dosya.
   - Banka istihbaratı ve piyasa görüşmeleri: `.txt`, `.pdf` veya `.docx`.
   - KKB/memzuç risk tablosu.
   - Olumsuz kayıtlar tablosu: karşılıksız çek, protesto, icra.
2. **Kontrol eder (kod):**

   | Kontrol | Önem |
   |---|---|
   | VKN kontrol hanesi; firma yaşı (3 yıldan genç) | yüksek / orta |
   | Ortaklık paylarının toplamı %100 değil | orta |
   | Tüzel kişi ortak (grup riski) | bilgi |
   | Son 12 ayda sicil değişiklikleri (sermaye, yönetim, adres, unvan); 3 ve üzeri | orta |
   | KKB'de görünüp istihbaratı alınmamış banka | orta |
   | İstihbarattaki nakdi risk ile KKB'nin farkı | bilgi |
   | Sektörde limit doluluğu ≥ %90; banka bazında ≥ %95 | orta / bilgi |
   | Olumsuz kayıt: açık veya son 1 yıl | yüksek |
   | Olumsuz kayıt: son 3 yıl | orta |
   | Olumsuz kayıt: daha eski | bilgi |
   | Metinlerde "gecikme", "sarkma", "icra", "protesto" gibi ifadeler | orta / yüksek |

   "Bulunmamaktadır", "duyulmadı" gibi olumsuzlanmış cümleler sayılmaz.
3. **Maskeler:**
   - Firma unvanı `[FİRMA]`, VKN `[VKN]` olur.
   - Sicildeki ortak ve yöneticiler otomatik olarak `[KİŞİ-n]` ve `[ŞİRKET-n]` olur; soyadları da maskelenir.
   - Telefon, e-posta, TCKN ve IBAN maskelenir.
   - Metinde geçen başka kişileri (eski yöneticiler, görüşülen kişiler) `--gizle` ile verin.
   - Veri gönderilmeden önce onayınız alınır.
4. **Yazar (model):**
   - Bölümler: künye, ortaklık ve yönetim, faaliyet ve piyasa, banka ilişkileri, olumsuz bilgiler.
   - Listeler: çelişkiler, güçlü yönler, risk işaretleri, teyit edilecekler.
   - Tarafsız bir genel değerlendirme; kredi kararı içermez.
   - Piyasa duyumlarını kesin bilgi gibi yazmaması istenir.
   - Metindeki, girdilerde olmayan sayılar **işaretlenir**.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                        # örnek: kurgusal metal işleme firması
python agent.py --girdi ./istihbarat_klasoru --tarih 08.10.2026
python agent.py --girdi ./istihbarat --gizle "Eski Yönetici Adı" "Görüşülen Kişi"
```

| Dosya | Biçim |
|---|---|
| `ticaret_sicil.txt` | `Unvan:`, `VKN:`, `Kuruluş Tarihi:`, `Sermaye:`, `Faaliyet Konusu:`, ardından `Ortaklar:` (`- Ad: %pay`), `Yönetim Kurulu:` (`- Ad (görev)`), `Değişiklikler:` (`- GG.AA.YYYY: açıklama`) |
| `banka_istihbarati.txt` | Satır başına `Banka Adı: metin` ("Nakdi limit … TL, risk … TL" ifadeleri okunur) |
| `piyasa_gorusmeleri.txt` | Serbest metin |
| `kkb_risk.csv` | Banka, Nakdi Limit, Nakdi Risk, Gayrinakdi Limit, Gayrinakdi Risk |
| `olumsuz_kayitlar.csv` | Tür, Tarih, Tutar, Durum (Açık / Ödendi), Açıklama |

## Çıktı

- **Rapor:** `istihbarat_raporu.md`, takma adlar geri açılmış ve tablolar koddan.
- **Excel:** `istihbarat_raporu.xlsx`:
  - `Özet`.
  - `Kontroller`: kod ve model bulguları, "İnceleme" sütunu.
  - `Ortaklık ve Sicil`.
  - `Banka İstihbaratı`: istihbarat ve KKB yan yana.
  - `KKB Risk`: "İstihbarat Var mı?" sütunuyla.
  - `Olumsuz Kayıtlar`.

## Dikkat

- **Taslak niteliği:** Rapor bir derleme taslağıdır. Bilgilerin güncelliği ve doğruluğu kaynaklara bağlıdır.
  Ticaret Sicili Gazetesi ve KKB sorgularını güncel olarak teyit edin.
- **Duyumlar:** Piyasa görüşmeleri duyum niteliğindedir. Kişiler hakkındaki bilgiler yalnız mevzuatın izin
  verdiği amaçla ve yetkili kişilerce kullanılmalıdır (KVKK, bankacılık sırrı).
- **Örnek veri:** Örnek firma, kişiler ve bankalar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Sicil ayrıştırma.
- Örnekteki bulgular: %95 ortaklık toplamı, istihbaratı alınmamış Bank D, açık icra, Bank B'de risk farkı.
- Kişi ve şirket adlarının modele gitmemesi.
- Uydurma bir sayının (4.800.000 TL) yakalanması.

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

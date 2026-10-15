# Şikâyet Analizi · AI Agent

> Müşteri Hizmetleri ve Çağrı Merkezi › Müşteri Deneyimi Uzmanı · Workers / Workless

Şikâyet kayıtlarını **konu, ürün ve olası kök neden** başlıklarına kümeler ve **eğilim raporu** çıkarır.
- **Kod:** Sayıları hesaplar: dönem bazında adetler, yükselen konular, Pareto, çözüm süresi, hedef süre aşımı,
  tekrarlayan şikâyetler.
- **Model:** Konu başlıklarını önerir ve her kaydı sınıflandırır: konu, ürün, olası kök neden, risk işaretleri.
- **Siz:** Sınıflandırmayı `Kontrol / Düzeltme` sütunuyla gözden geçirirsiniz.

Herhangi bir sektörde kullanılabilir: perakende, e-ticaret, banka, sigorta, telekom, üretim, hizmet.

## Nasıl çalışır?

1. **Okur ve hesaplar (kod):**
   - Başlık satırını kendisi bulur; üstteki rapor başlıklarını atlar. Sütun adları esnektir ("Şikâyet",
     "Açıklama", "Metin"...).
   - Çözüm süresini (kapanış − açılış) ve açık kayıtların yaşını hesaplar; **hedef süreyi** (`--hedef-gun`,
     varsayılan 15 takvim günü) aşanları işaretler.
   - Aynı müşterinin 30 gün içindeki **tekrarlayan** şikâyetlerini bağlar.
2. **Maskeler:** Müşteri adı metinde `[MÜŞTERİ]` olur. `--gizle` ile verilen adlar `[GİZLİ-n]` olur. Telefon,
   e-posta, TCKN ve IBAN maskelenir. Müşteri numarası gönderilmez. Veri gönderilmeden önce onayınız alınır.
3. **Kümeler (model):** `--konular` verilmezse model kayıtlardan bir örneklemle 5-12 konu başlığı önerir:
   "Teslimat gecikmesi", "Para iadesi gecikmesi" gibi. Öneri `Konu Listesi` sayfasına yazılır.
4. **Sınıflandırır (model):** Her kayıt için şunları çıkarır:
   - Tek bir konu.
   - Ürün veya hizmet.
   - Olası kök neden kategorisi: ürün kusuru, teslimat, fiyat/fatura, personel, süreç, bilgilendirme, sistem,
     tedarikçi/iş ortağı, müşteri beklentisi ya da belirsiz.
   - Metinde açıkça geçen risk işaretleri: hakem heyeti/dava, resmî kurum başvurusu, sosyal medya,
     sağlık-güvenlik, kişisel veri, kaba davranış.

   Listede olmayan konu "Diğer" sayılır ve raporda belirtilir.
5. **Eğilim (kod):**
   - Konu × dönem (ay veya hafta) tablosu.
   - Son dönem önceki üç dönemin ortalamasının 1,5 katını aşıyor ve en az `--esik` (3) kayıtsa konu
     **Yükselen** olur; yarısının altına düşmüşse **Azalan**.
   - Son dönem eksikse (ör. ay bitmeden) uyarı verilir.
   - Pareto ile şikâyetlerin %80'ini oluşturan konular "A" olarak işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek: 38 şikâyet, Temmuz-Eylül 2026
python agent.py --girdi sikayetler.xlsx --konular konular.txt     # sabit konu listesiyle
python agent.py --girdi sikayetler.xlsx --donem hafta --hedef-gun 10 --tarih 30.09.2026
```

| Dosya | Biçim |
|---|---|
| Şikâyet kayıtları (`.xlsx` / `.csv`) | `Şikâyet` (metin) zorunlu. İsteğe bağlı: `Şikâyet No`, `Tarih`, `Kanal`, `Müşteri No`, `Müşteri Adı`, `Ürün`, `Durum`, `Kapanış Tarihi` |
| `konular.txt` (isteğe bağlı) | Satır başına `Konu: tanım`. `#` ile başlayan satırlar yok sayılır. |

**İpucu:** İlk çalıştırmada modelin önerdiği konuları `Konu Listesi` sayfasından alın, düzeltin ve sonraki
dönemlerde `--konular` ile verin. Böylece her ay aynı başlıklar kullanılır ve eğilim karşılaştırılabilir olur.

## Çıktı

`sikayet_analizi.xlsx`:
- `Özet`: göstergeler, yükselen ve azalan konular, uyarılar.
- `Konu Analizi`: Pareto, ortalama çözüm süresi, hedef aşımı, eğilim, grafik.
- `Eğilim`: konu × dönem tablosu ve ilk 5 konunun grafiği.
- `Kök Neden`: konu × olası kök neden ve ürün × konu tabloları.
- `Takip Listesi`: risk işaretli, hedef süreyi aşan ve tekrarlayan açık kayıtlar; "Yapılacak / Sorumlu" sütunu.
- `Kayıtlar`: tüm kayıtlar ve model sonuçları (mor), "Kontrol / Düzeltme" sütunu.
- `Konu Listesi`: kullanılan konular ve tanımları.

## Dikkat

- **Kök neden:** Model yalnız şikâyet metnine bakar. Bulduğu kök neden bir **öneridir**. Kesin kök neden
  analizi (5 Neden, balık kılçığı) ilgili ekiplerle ve iç verilerle yapılmalıdır.
- **Hedef süre:** 15 gün yalnız örnektir. Kendi şikâyet yönetimi prosedürünüzü veya sektörünüzün mevzuatında
  tanımlı süreyi kullanın. Şikâyet yönetimi için TS ISO 10002 kılavuz olarak kullanılabilir.
- **Risk işaretleri:** Yalnız metinde açıkça geçenler işaretlenir. Risk işareti olmayan bir kaydın risksiz
  olduğu anlamına gelmez.
- **Kişisel veri:** Ad maskelemesi yardımcıdır, garanti değildir. Metinde başka kişilerin adları geçebilir;
  bunları `--gizle` ile verin veya yerel model (`WW_PROVIDER=ollama`) kullanın.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Hedef süre aşımı, tekrarlayan şikâyet ve başlık satırı tespiti.
- Maskeleme: ad ve telefon modele gitmez.
- Eğilim kuralları; örnekte teslimat gecikmesi Eylül'de 2'den 7'ye çıkar ve "Yükselen" olur.
- Model listede olmayan bir konu verirse kaydın "Diğer" sayılması.

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

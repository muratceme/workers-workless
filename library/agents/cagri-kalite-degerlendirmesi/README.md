# Çağrı Kalite Değerlendirmesi · AI Agent

> Müşteri Hizmetleri ve Çağrı Merkezi › Müşteri Deneyimi Uzmanı · Workers / Workless

Görüşme dökümünü **kendi kalite formunuzdaki** kriterlere göre değerlendirir. Her kriter için sonuç, gerekçe ve
dökümden birebir alıntı verir; puanı hesaplar, kritik hataları işaretler ve temsilciye verilecek **geri bildirim
notunu** taslak olarak yazar. Alıntısı doğrulanamayan kriterler puana katılmaz ve insan incelemesine ayrılır.

## Nasıl çalışır?

1. **Kod (hazırlık):**
   - Dökümü okur: `Temsilci:` ve `Müşteri:` ile başlayan satırlar; baştaki zaman damgaları atılır. Dosyanın
     başındaki `Temsilci Adı:`, `Tarih:`, `Çağrı No:` satırları bilgi olarak alınır.
   - Telefon, e-posta, TCKN, IBAN ve kart numaralarını maskeler.
2. **Model (değerlendirme):** Her kriter için `karşılandı`, `kısmen`, `karşılanmadı` veya `uygulanamaz` sonucunu,
   gerekçeyi ve temsilcinin sözlerinden birebir bir alıntıyı verir. Güçlü yönleri, gelişim alanlarını ve geri
   bildirim notunu yazar.
3. **Kod (doğrulama ve puan):**

   | Durum | Sonuç |
   |---|---|
   | "Karşılandı" / "kısmen" ama alıntı temsilci sözlerinde birebir yok | İnceleme (insan) |
   | Alıntı yalnız müşterinin sözlerinde geçiyor | İnceleme (insan) |
   | Model "karşılandı" dedi ama formdaki anahtar ifadelerden hiçbiri temsilci sözlerinde yok | İnceleme (insan) |
   | Türü "İhlal" olan kriterde "karşılanmadı" ama ihlali gösteren alıntı yok | İnceleme (insan) |
   | Model kriter için sonuç döndürmedi | Değerlendirilmedi |
   | Formda olmayan kriter kodu | Atılır |

   - **Puan** = Σ ağırlık × (karşılandı 1; kısmen 0,5; karşılanmadı 0) / Σ ağırlık × 100. "Uygulanamaz", "inceleme" ve
     "değerlendirilmedi" kriterleri paya ve paydaya girmez.
   - **Kritik hata:** Kritik bir kriter karşılanmadıysa toplam puan 0 olur; ham puan ayrıca gösterilir.
     `--kritik-sifirlama-yok` ile puan sıfırlanmaz, yalnız işaretlenir.
   - Kritik kriter incelemeye düştüyse yüksek önemli uyarı verilir.

### Kalite formu

| Sütun | Açıklama |
|---|---|
| Kod, Bölüm, Kriter, Açıklama | Açıklama modelin kriteri doğru anlaması için önemlidir; beklenen davranışı yazın |
| Puan | Kriterin ağırlığı |
| Kritik | Evet / Hayır |
| Anahtar İfadeler | İsteğe bağlı; `|` ile ayrılmış. Kriter ancak bu ifadelerden biri temsilci sözlerinde geçiyorsa "karşılandı" kalır (ör. `talep numara|kayıt numara`) |
| Tür | `Davranış` (yapılması gereken; varsayılan) veya `İhlal` (yapılmaması gereken, ör. yasaklı ifade, gereksiz veri isteme). İhlal türünde "karşılandı" için alıntı aranmaz; "karşılanmadı" için ihlali gösteren alıntı aranır |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek: 2 görüşme, 10 kriterlik form
python agent.py --dokum gorusmeler/ --form kalite_formu.xlsx
python agent.py --dokum gorusme_0412.txt --form kalite_formu.csv --kritik-sifirlama-yok
```

Döküm biçimi:

```text
Temsilci Adı: Örnek Temsilci
Tarih: 05.10.2026
Temsilci: Örnek İletişim'e hoş geldiniz, size nasıl yardımcı olabilirim?
Müşteri: Faturam hakkında bilgi almak istiyorum.
```

## Çıktı

`cagri_kalite_degerlendirmesi.xlsx`:
- `Özet`: görüşme başına puan, ham puan, kritik hata, inceleme bekleyen kriter sayısı, temsilcinin konuşma payı; boş
  "Değerlendiren" sütunu.
- `Kriter Sonuçları`: sonuç, alınan puan, gerekçe, doğrulanmış alıntı, kontrol notu; boş "Uzman Kararı" sütunu.
- `Geri Bildirim`: güçlü yönler, gelişim alanları, taslak not.
- `Uyarılar`.

Mor hücreler yapay zekâ tarafından yazılmıştır.

## Dikkat

- **Sonuçlar taslaktır.** Puanlar performans değerlendirmesinde, primde veya disiplin sürecinde kullanılmadan önce
  bir kalite uzmanı kaydı dinleyerek onaylamalıdır. Çalışan hakkında yalnız otomatik değerlendirmeye dayanan karar
  verilmemelidir (KVKK md. 11).
- **Döküm sınırı:** Model yalnız yazıya dökülmüş sözleri görür. Ses tonu, bekleme süresi, sessizlik ve ekranda
  yapılan işlemler değerlendirilemez. Otomatik yazıya dökme hataları sonucu etkiler.
- **Kişisel veri:** Maskeleme yardımcıdır, garanti değildir; adlar ve adresler maskelenmez. Görüşme kayıtlarının
  işlenmesi ve API'ye gönderilmesi için aydınlatma ve yurt dışına aktarım şartlarını kontrol edin. Yerel model için
  `WW_PROVIDER=ollama` kullanabilirsiniz.
- **Örnek veri:** Örnek görüşmeler, firma ve kişiler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Form (ağırlık, kritik, ihlal türü) ve döküm okuma; maskeleme (telefon, kart).
- Alıntının temsilci / müşteri sözlerinde aranması.
- İnceleme kuralları: metinde olmayan alıntı, müşteri sözü, anahtar ifade eksikliği.
- Puan hesabı (elle), kritik hata ve sıfırlama seçeneği, formda olmayan kod, Excel ve CLI.

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

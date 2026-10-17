# Etiket Bilgisi Kontrolü · AI Agent

> Kalite Güvence ve Gıda Güvenliği › Gıda Mühendisi · Gerisi Bende

Gıda etiketini **Türk Gıda Kodeksi Gıda Etiketleme ve Tüketicileri Bilgilendirme Yönetmeliği**'ne göre ön kontrolden
geçirir. Zorunlu bilgiler, temel görüş alanı, yazı büyüklüğü, net miktar birimi ve beslenme bildirimini **kod**
kontrol eder. Bileşen listesindeki alerjenleri, miktarı belirtilmesi gereken bileşenleri ve yanıltıcı olabilecek
ifadeleri **model** yorumlar. Kod, modelin bulgularını bileşen listesiyle karşılaştırır ve vurgulanmamış alerjenleri
işaretler.

## Nasıl çalışır?

1. **Biçimsel kontrol (kod):**
   - Zorunlu bilgiler (md. 9): ad, bileşenler, net miktar, TETT / SKT, muhafaza koşulları, işletmeci adı ve adresi,
     işletme kayıt / onay numarası, beslenme bildirimi; menşe ve parti numarası "kontrol edilmeli"
   - Ürün adı, net miktar (ve varsa alkol) temel görüş alanında mı (md. 16/4)
   - x-yüksekliği en az 1,2 mm; en geniş yüzeyi 80 cm²'den küçük ambalajda en az 0,9 mm (md. 16)
   - Net miktar birimi: katıda g / kg, sıvıda ml / L (md. 26)
   - Beslenme bildirimi: enerji (kJ ve kcal), yağ, doymuş yağ, karbonhidrat, şeker, protein, tuz bu sırayla
   - "Katkısız" ifadesi varken bileşenlerde katkı maddesi sınıfı ya da E kodu
2. **Yorum (model):** Alerjen içeren bileşenler (Ek-1'deki 14 grup), adında vurgulanan bileşenler (QUID), yanıltıcı
   olabilecek ifadeler, diğer notlar.
3. **Denetim (kod):** Aşağıdakilerden biri varsa sonuç modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alerjen bileşeni listede aynen geçmiyor; kodun anahtar kelimeyle bulduğu bir alerjen grubu (ör. "peynir altı suyu"
     → süt) atlanmış
   - QUID bileşeni listede yok ya da ürün adında vurgulanmıyor; "yanıltıcı" denen ifade etikette yok
   - "Katkısız" çelişkisi değerlendirilmemiş
4. **Sonuç (kod):** Her alerjenin bileşen listesinde vurgulanıp vurgulanmadığı (BÜYÜK HARF ya da `**…**`) ve QUID
   yüzdesinin yazılıp yazılmadığı belirlenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                # örnek (kurgusal) çilekli kek etiketi
python agent.py --etiket etiket.docx
```

Etiket metnini ambalajdan aktarırken vurgulu yazılmış alerjenleri BÜYÜK HARF ya da `**yıldız**` arasında yazın; ön yüzde
yer alan bilgileri "Temel Görüş Alanı:" satırına, ölçtüğünüz küçük harf yüksekliğini "x-Yüksekliği:" satırına girin.

## Çıktı

- `cikti/etiket_kontrol.xlsx`: Kontrol (kırmızı: uygunsuz, sarı: kontrol edilmeli, yeşil: uygun), Bileşenler.

## Dikkat

- **Ön kontroldür.** Ürüne özgü kodeks tebliğleri (ör. süt ürünleri, et ürünleri, takviye edici gıdalar), beslenme ve
  sağlık beyanları, net miktar yazı yüksekliği ve tarih yazım biçimleri bu agent'ta kontrol edilmez. Etiket, yönetmeliğin
  ve Tarım ve Orman Bakanlığı kılavuzlarının güncel hâline göre gıda mühendisi tarafından onaylanmalıdır.
- **Çapraz bulaşma** ("… içerebilir") uyarısı tesisinizdeki alerjen yönetimine göre ayrıca değerlendirilir.
- **Örnek veri:** Ürün ve işletme kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk çıktıda listede olmayan "buğday nişastası"nı yazar, süt ve soyayı atlar,
adda geçmeyen "kakao" için QUID ister, etikette olmayan "Glutensiz" ifadesini yanıltıcı sayar ve "Katkısız" çelişkisini
atlar. Kodun altı sorunu da yakaladığı; x-yüksekliği (1,0 < 1,2 mm), katkısız çelişkisi ve vurgulanmamış üç alerjen
(buğday unu, peynir altı suyu tozu, soya lesitini) bulgularının raporda yer aldığı doğrulanır.

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

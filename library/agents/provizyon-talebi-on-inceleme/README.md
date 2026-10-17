# Provizyon Talebi Ön İnceleme · AI Agent

> Sağlık Sigortası › Provizyon Uzmanı · Gerisi Bende

Anlaşmalı kurumlardan gelen **provizyon taleplerini** (talep formu, epikriz, hekim notu) poliçe bilgileri ve sizin
verdiğiniz özel / genel şartlar metniyle **ön incelemeden** geçirir. **Karar vermez.** Onay, ret ve tıbbi değerlendirme
provizyon uzmanı ile hekime aittir. Agent işaretleri, ilgili maddeleri ve ön hesabı hazırlar.

## Nasıl çalışır?

1. **Hazırlık (kod):** Talep belgeleri, poliçe tablosu ve şartlar metni okunur. "Hasta:" satırındaki ad ve `--gizle`
   adları `[HASTA-n]` olur; TCKN, telefon ve e-posta maskelenir. Belgede geçen poliçenin başlangıç tarihi ve planı
   modele verilir.
2. **İnceleme (model):**
   - talep bilgileri: poliçe no, yatış türü ve tarihi, gün, tanı, ICD-10, işlem, talep tutarı;
   - poliçe öncesi var olabilecek hastalık işaretleri (belgeden alıntıyla);
   - ilgili şart maddeleri (şartlar metninden alıntıyla);
   - eksik bilgiler ve öneri: Onay önerilir / Tıbbi inceleme gerekir / Ek bilgi iste.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o talep modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı belgede ya da şartlar metninde aynen geçmiyor; madde şartlar metninde yok
   - Tarih, tutar, gün ya da ICD-10 kodu belgede yazmıyor (uydurma ICD kodu kabul edilmez); ICD biçimi hatalı
   - Poliçe tabloda yok
   - İşaret varken "Onay önerilir" seçilmiş
4. **Poliçe kontrolleri (kod):**

| Kontrol | Kural |
|---|---|
| Bekleme süresi | Planlı yatışta başlangıç + bekleme süresi > yatış tarihi → Dikkat (acil ve ayakta hariç) |
| Poliçe başlangıcı | İşlem tarihi başlangıçtan önce → Hata |
| Limit | Talep > yatarak / ayakta limit → Dikkat; ön hesap limitten yapılır |
| Katılım payı | Esas tutar × katılım %; sigorta payı = esas − katılım |

**Durum:**
- **Eksik bilgi:** kod "Hata" bulduysa ya da model ek bilgi istediyse.
- **İnceleme:** kod bulgusu, işaret ya da ilgili madde varsa.
- **Uygun:** yukarıdakilerin hiçbiri yoksa.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek 3 talep (1 inceleme, 2 uygun)
python agent.py --talepler ./talepler --policeler policeler.xlsx --sartlar ozel_sartlar.txt --gizle "Ayşe Deneme"
```

Şartlar metnine poliçenizin özel / genel şartlarından ilgili maddeleri kopyalayın; agent şartları kendisi bilmez.

## Çıktı

`cikti/provizyon_on_inceleme.xlsx`:
- `Talepler`: talep başına bir satır. İçerik:
  - talep bilgileri;
  - limit, katılım payı ve sigorta payı ön hesabı;
  - kod bulguları;
  - önceden var olabilecek hastalık işaretleri (“alıntı” — açıklama);
  - ilgili maddeler, eksik bilgiler;
  - model önerisi ve gerekçesi;
  - durum (Uygun / İnceleme / Eksik bilgi).
- `Kontrol`

## Dikkat

- **Ön incelemedir; provizyon kararı değildir.** "Önceden var olabilecek hastalık" işaretleri yalnız belgedeki
  ifadelerdir, tıbbi değerlendirme değildir. Ret kararları poliçe şartlarına göre hekim ve yetkili uzman tarafından
  verilmeli ve gerekçelendirilmelidir.
- **Ön hesap** anlaşmalı kurum fiyatlarını, provizyon sonrası fatura kontrolünü ve özel limitleri (oda, ilaç) kapsamaz.
- **Sağlık verisi özel nitelikli kişisel veridir (KVKK md. 6).** Tanı ve hekim notları modele gider. Mümkünse yerel model
  (Ollama) kullanın; şirket politikanızı kontrol edin.
- **Örnek veri:** Poliçeler, şartlar, hastalar ve talepler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model bel fıtığı talebinde "4 yıldır bel ağrısı" işaretini bulduğu hâlde ilk turda
bilerek "Onay önerilir" der. Kodun bunu yakalayıp yalnız o talebi yeniden yazdırdığı doğrulanır. Ayrıca şunlar
doğrulanır:
- bekleme süresi, başlangıç öncesi işlem, limit, katılım ve sigorta payı hesapları,
- uydurma ICD kodu, olmayan madde ve belgede olmayan alıntı yakalama,
- hasta adlarının hiçbir mesajda geçmemesi,
- poliçe başlangıç tarihinin modele verilmesi.

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

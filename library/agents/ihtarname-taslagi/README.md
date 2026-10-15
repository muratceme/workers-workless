# İhtarname Taslağı · AI Agent

> Hukuk › Avukat · Workers / Workless

Olay özeti ve talepten **noter ihtarnamesi taslağı** hazırlar. **Taraf adları, adresleri ve kimlik numaraları yapay zekâ
modeline gönderilmez;** olay metnindeki adlar `{KESIDECI}` / `{MUHATAP}` ile değiştirilir, taslağa kod geri yazar.

## Nasıl çalışır?

1. **Hazırlık (kod):** Olay özeti satır satır numaralanır, taraf bilgileri ayrı dosyadan okunur.
2. **Taslak (model):** Konu, açıklamalar (her paragraf dayandığı satırlarla) ve sonuç-istem yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Paragrafta dayandığı satırlarda olmayan tutar, tarih ya da sayı var
   - Kanun maddesine atıf yapılmış ama madde olayda, talepte ya da `--dayanak` listesinde yok
   - Sonuç-istemde verilen süre ("30 gün") yazılmamış ya da olayda olmayan sayı eklenmiş
   - Kira temerrüdü ihtarında süre içinde ödenmezse sözleşmenin feshedileceği yazılmamış
   - `{KESIDECI}` / `{MUHATAP}` kullanılmamış, yer tutucuya ek getirilmiş ya da hakaret sayılabilecek ifade var
4. **Belge (kod):** Noterlik başlığı, keşideci / vekil / muhatap bilgileri ve noterlik talep kalıbı ("bir nüshasının
   muhataba tebliğini, bir nüshasının dairenizde saklanmasını ve tebliğ şerhli bir nüshasının tarafıma verilmesini…")
   eklenir.

**Kira temerrüdü:** `--tur-ihtar kira-konut` ya da `kira-diger` seçilirse süre TBK m.315'e göre denetlenir. Konut ve çatılı
işyeri kiralarında en az 30 gün, diğer kiralarda en az 10 gün verilmezse agent çalışmaz.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                                  # örnek (kurgusal) kira alacağı
python agent.py --olay olay.txt --taraflar taraflar.txt --talep "…" --sure 30 --tur-ihtar kira-konut
python agent.py --olay olay.docx --taraflar taraflar.txt --talep "Ayıplı malın değiştirilmesi" --sure 7 --tur-ihtar genel --dayanak "TBK 227"
```

Taraf dosyası örneği:

```
Noterlik: İstanbul ( ) Noterliğine
Keşideci: Ad Soyad / Unvan
Keşideci Adres: …
Vekil: Av. Ad Soyad          (isteğe bağlı)
Muhatap: Ad Soyad / Unvan
Muhatap Adres: …
Muhatap Kimlik No: …         (isteğe bağlı)
```

## Çıktı

- `cikti/ihtarname_taslagi.txt`: notere verilecek biçimde taslak (imza alanı boş, "TASLAKTIR" notuyla).
- `cikti/ihtarname_taslagi.xlsx`:
  - `Paragraflar`: her paragraf ve dayandığı olay satırları.
  - `Eksik Bilgiler`: teyit edilmesi gereken bilgiler ve avukata iç not.

## Dikkat

- **Taslaktır.** Hukuki nitelendirme, süreler, talepler ve sonuçları avukat tarafından kontrol edilmelidir.
- **Süre ve tebliğ.** Süreler genellikle ihtarnamenin tebliğini izleyen günden işler. Özel süreler (TBK m.315 dışındakiler)
  agent tarafından denetlenmez; `--sure` değerini olaya göre siz belirleyin.
- **Kişisel veri.** Olay metninde taraf adı dışında kalan kişisel veriler (üçüncü kişi adları vb.) modele gidebilir;
  telefon, e-posta ve T.C. kimlik numarası maskelenir. Gerekirse olay metnini göndermeden önce sadeleştirin.
- **Örnek veri:** Taraflar ve olay kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek tutarı 80.000 TL yazar, olayda olmayan "TBK 299" maddesine
atıf yapar, "dolandırıcı" der, yer tutucuya ek getirir, süreyi 15 gün yazar ve fesih ihtarını unutur. Kodun bunları
yakalayıp düzelttirdiği, taraf adlarının modele gönderilmediği ve TBK 315'e aykırı sürede çalışmadığı doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6098 sayılı Türk Borçlar Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6098.pdf) | Resmî metinle karşılaştırılan maddeler: 100, 115, 182, 315 |

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

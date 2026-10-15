# Müşteri İhtiyaç Analizi · AI Agent

> Acente Teknik › Sigorta Uzmanı (Acente) · Workers / Workless

Müşteriyle yapılan görüşmenin notlarından **risklerini ve acentenizin ürün kataloğundan önerilecek sigorta ürünlerini**
içeren ihtiyaç analizi hazırlar. Öneriler yalnız kataloğunuzdaki ürünlerden yapılır; prim ve teminat tutarı uydurulmaz.

## Nasıl çalışır?

1. **Hazırlık (kod):** Görüşme notu satır satır numaralanır; telefon, e-posta ve kimlik numarası maskelenir. Ürün
   kataloğu okunur (Ürün; Branş; Kapsam; Zorunlu).
2. **Analiz (model):** Müşteri profili, riskler (öncelik ve dayanak satırlarıyla), öneriler (karşıladığı risklerle),
   eksik bilgiler ve müşteriye sorulacak sorular yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa analiz modele geri gönderilir (varsayılan en fazla 2 tur):
   - Önerilen ürün katalogda yok
   - Risk ya da önerinin dayanak satırı geçersiz
   - Bir risk hiçbir öneriyle karşılanmamış (katalogda karşılığı yoksa "katalog dışı" işaretlenmeli)
   - Zorunlu bir sigorta ele alınmamış: notlarda araç geçiyorsa zorunlu trafik, kendi konutu geçiyorsa DASK
     (katalogda "Zorunlu: E" olan ürünler)
   - Notlarda ve katalogda olmayan sayı (prim, teminat, limit) yazılmış
   - "Tam koruma", "her şeyi karşılar", "garanti" gibi aşırı vaat var

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                  # örnek görüşme + örnek katalog
python agent.py --girdi gorusme.txt --katalog urunler.xlsx
```

Katalogda "Branş" sütunu zorunlu sigorta denetimi için kullanılır: `Oto` branşındaki zorunlu ürün araç geçen notlarda,
`Konut` branşındaki zorunlu ürün kendi konutu geçen notlarda aranır.

## Çıktı

- `cikti/ihtiyac_analizi.xlsx`:
  - `Riskler`: öncelik, katalog dışı işareti, karşılayan ürünler ve dayanak satırlar.
  - `Öneriler`: ürün, branş, durum (yeni / mevcut), gerekçe, zorunlu işareti.
  - `Sorular`: eksik bilgiler ve müşteriye sorulacaklar.
- `cikti/ihtiyac_analizi.txt`: müşteriyle paylaşılabilecek ön değerlendirme özeti.

## Dikkat

- **Ön analizdir, teklif değildir.** Teminat, limit, muafiyet ve istisnalar sigorta şirketinin genel ve özel şartlarına
  göre belirlenir. Müşteriye önerilen ürünlerle ilgili bilgilendirme yükümlülüklerinizi (bilgilendirme formu vb.) yerine
  getirin.
- **Zorunlu sigortalar.** Kod yalnız trafik ve DASK için anahtar kelime denetimi yapar. Mesleki sorumluluk gibi
  faaliyete bağlı diğer zorunlu sigortaları siz değerlendirin.
- **Kişisel veri.** Görüşme notları (yaş, aile durumu, sağlık tercihi gibi bilgiler) modele gönderilir; telefon, e-posta ve
  kimlik numarası maskelenir. Müşteriden KVKK kapsamında gerekli aydınlatma ve onayı alın; notlara ad soyad yazmayın.
- **Örnek veri:** Müşteri ve katalog kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk analizde bilerek:
- katalogda olmayan "Kasko Plus" ürününü önerir,
- kendi konutunda oturan müşteriye DASK'ı atlar (deprem riski karşılanmaz),
- "aylık 850 TL primle tam koruma sağlar" yazar.

Kodun bunları yakalayıp düzelttirdiği doğrulanır.

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

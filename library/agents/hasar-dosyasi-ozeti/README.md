# Hasar Dosyası Özeti · AI Agent

> Sigorta › Hasar Yönetimi › Hasar Uzmanı · Gerisi Bende

Bir hasar dosyasındaki belgeleri okuyup **karar için tek sayfalık özet** hazırlar. Belgeler poliçe özeti, ihbar
ve beyan, kaza tespit tutanağı, ekspertiz raporu ve faturalar olabilir. Belgeler arası tutarsızlıkları
**kodla** bulur; değerlendirmeyi model yazar, kararı hasar uzmanı verir.

## Nasıl çalışır?

1. **Belgeleri ayırır (kod):** Klasördeki `.pdf`, `.docx` ve `.txt` dosyaları okunur ve türüne göre ayrılır:
   poliçe, ihbar/beyan, tutanak, ekspertiz, fatura, ehliyet, ruhsat.
2. **Kontrol eder (kod):** Bu bulgular kesindir ve modele "değiştirme" talimatıyla verilir.
   - Belgelerde farklı yazılmış **hasar tarihi**
   - Sigortalı plakasına benzeyen ama farklı **plaka** (ör. faturada `34 ABC 128`)
   - **Poliçe dönemi dışında** hasar ve poliçe başlangıcına yakın (30 gün) hasar
   - Fatura KDV hariç tutarının **eksper onaylı tutarı aşması**
   - İhbar gecikmesi ve **eksik evrak** (varsayılan liste veya `--evrak` dosyası)
3. **Maskeler:** Plakalar her belgede aynı takma adla (`[PLAKA-A]`) değiştirilir. `--gizle` ile verilen kişi
   adları `[GİZLİ-1]` olur. Telefon, e-posta, TCKN ve IBAN maskelenir. Veri gönderilmeden önce onayınız alınır.
4. **Özetler (model):** Olay, kusur ve rücu, teminat değerlendirmesi (taslak), tutarsızlıklar (kaynak belge
   adlarıyla), açık sorular, önerilen adımlar ve karar özeti yazılır. Modelin uydurduğu belge adları
   "kaynak doğrulanamadı" olarak işaretlenir.
5. **Rapor:** Tek sayfalık `.md` özet ve Excel üretilir. Rapor şirket içinde kaldığı için takma adlar
   gerçek değerlere geri çevrilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                          # örnek kasko dosyasıyla
python agent.py --girdi ./dosya_2026_0457 --gizle "Sigortalı Adı" "Karşı Sürücü"
python agent.py --girdi ./dosya --evrak evrak_listesi.txt                # branşınıza göre evrak listesi
```

**Evrak listesi:** Her satıra bir belge türü yazılır. Kullanılabilecek türler: Poliçe, İhbar / beyan, Kaza
tespit tutanağı, Ekspertiz raporu, Fatura, Ehliyet, Ruhsat.

Fotoğraflar ve taranmış PDF'ler okunmaz; bunlar "okunamayan" olarak listelenir.

## Çıktı

`hasar_ozeti.md` (tek sayfa) · `hasar_ozeti.xlsx`: `Özet` · `Kontroller` (önem, bulgu, kaynak, 'Uzman Onayı') · `Belgeler`

## Testler

Testler gerçek API çağırmaz. Örnek dosyada bilerek bırakılmış dört tutarsızlığın ve iki eksik evrakın kodla
bulunduğu doğrulanır.

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

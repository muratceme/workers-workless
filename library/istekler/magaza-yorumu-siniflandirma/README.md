# Mağaza Yorumu Sınıflandırma · AI Agent

> Sizden Gelenler › Abonelikli mobil uygulama · Gerisi Bende

Google Play ve App Store yorumlarını **kategori ve duyguya** ayırır, **özellik taleplerini gruplar** (en çok
istenen ilk sırada), her yoruma geliştirici adına **cevap taslağı** yazar ve sürüm bazında puanı izler. Agent
mağazaya cevap göndermez; taslakları siz okuyup gönderirsiniz.

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):** Play Console yorum dökümünü (UTF-16 CSV) ve App Store yorumlarını okur. Metni olmayan
   yalnız puanlar modele gitmez, istatistiğe girer. E-posta ve telefon maskelenir.
2. **Model:** Kategori, duygu, varsa özellik talebinin kısa adı, yorumdan birebir kısa alıntı ve cevap taslağı
   yazar. Önceki paketlerde kullandığı talep adlarını görür; aynı isteğe aynı adı verir.
3. **Kod modeli denetler:**
   - 1–2 yıldızlı yorum "olumlu" sayılamaz; **olumsuz** yapılır.
   - Alıntı yorumda **birebir** geçmiyorsa silinir (model uydurmasın).
   - Cevap taslağı karakter sınırını aşıyorsa **kırmızı** işaretlenir.
   - Özellik talepleri aynı adla gruplanır; yorum sayısı ve ortalama puanla sıralanır.

## Yorumları nereden alırım?

- **Google Play:** Play Console'un *Raporları indir* bölümündeki yorum raporu ya da Google Cloud Storage'daki aylık
  dosya (`reviews_reviews_<paket>_<YYYYMM>.csv`). Dosya UTF-16 kodludur; paket bunu kendisi tanır.
- **App Store:** Yorumları App Store Connect API'si ya da kullandığınız araçla (n8n, Make vb.) `Tarih, Puan,
  Başlık, Yorum, Sürüm` sütunlarıyla bir CSV'ye yazın.

## Cevap sınırı

Varsayılan sınır **350 karakter**dir; `--cevap-siniri` ile değiştirin. Mağazaların geliştirici cevabı sınırlarını
güncel belgelerinden kontrol edin.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                           # örnek 24 yorum (kurgusal)
python agent.py --play reviews_reviews_com.sirket.uygulama_202609.csv
python agent.py --play play_yorumlari/ --appstore appstore_yorumlari.csv --urun "burç ve fal uygulaması"
```

## Çıktı

- `cikti/magaza_yorumlari.xlsx`: `Olumsuz` · `Özellik Talepleri` · `Yorumlar` (mor sütunlar yapay zekâ, "Onay"
  sütunu sizin) · `Sürümler` (sürüm başına ortalama puan, 1–2 yıldız oranı, en sık olumsuz konu) · `Özet`
- `cikti/magaza-yorumu-siniflandirma.ozet.json`: **Uygulama Yönetim Özeti** paketinin okuduğu günlük yorum, ortalama
  puan ve olumsuz sayısı ile en çok istenen özellikler.

## Testler

Testler gerçek API çağırmaz. Sahte model bilerek her yoruma "olumlu" der, bir yorumda uydurma alıntı ve bir yorumda
400 karakterlik cevap verir. Kodun bunları düzelttiği ya da işaretlediği ve talepleri doğru grupladığı doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 21.10.2026. Sütun adları aşağıdaki belgeden alınmıştır; Google biçimi değiştirebilir, kendi
dosyanızın başlık satırıyla karşılaştırın.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| Google Play Console: rapor türleri (yorum raporu sütunları, UTF-16) | [support.google.com](https://support.google.com/googleplay/android-developer/answer/6135870) | Sütun adları belgeyle karşılaştırıldı |

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

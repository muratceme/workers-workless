# Trend Raporu Özeti · AI Agent

> Tasarım › Moda Tasarımcısı · Gerisi Bende

Sezon trend raporlarını koleksiyonunuza **uygulanabilir renk, kumaş, silüet, detay ve desen maddelerine** indirir. Her
madde rapordan aynen alıntıyla gelir. Koleksiyon brifinize uymayan trendler gerekçesiyle ayrı listelenir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Trend raporu satır (paragraf) satır numaralanır. Koleksiyon brifinden ürün grupları ve
   kaçınılacaklar okunur. Raporda hangi kategorilerin geçtiği anahtar kelimelerle belirlenir.
2. **Özet (model):** Her trend için kategori, kısa ad, rapordan alıntı, koleksiyona uygulama, ürün grupları ve öncelik
   yazılır. Brife uymayan trendler "uygulanmayacaklar" bölümüne gider.
3. **Denetim (kod):** Aşağıdakilerden biri varsa özet modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı dayandığı rapor satırında aynen geçmiyor
   - Ürün grubu brifte yok
   - Pantone (TCX / TPX) kodu ya da sayı raporda / brifte yok (renk kodu uydurulamaz)
   - Raporda geçen bir kategoride hiç madde yok
   - Brifteki kaçınılacaklardan biri (ör. payet, deri) uygulama maddesinde önerilmiş

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                  # örnek (kurgusal) trend notları + brif
python agent.py --rapor trend_raporu.pdf --brif koleksiyon_brifi.txt
```

Brif örneği:

```
Sezon: İlkbahar/Yaz 2027
Hedef Müşteri: 25-40 yaş arası, şehirde çalışan kadın
Ürün Grupları: Gömlek, Pantolon, Elbise, Etek, Tişört
Fiyat Segmenti: Orta
Ana Kumaşlar: Pamuk, keten, viskon
Kaçınılacaklar: Payet, metalik kumaş, deri
```

## Çıktı

- `cikti/trend_ozeti.xlsx`:
  - `Maddeler`: kategori, trend, rapordan alıntı, satır, koleksiyona uygulama, ürün grupları, öncelik ve "Karar" sütunu.
  - `Uygulanmayacaklar`: gerekçesiyle.
- `cikti/trend_ozeti.txt`: kategori bazında özet.

## Dikkat

- **Lisanslı raporlar.** Ücretli trend servislerinin raporlarını bir yapay zekâ sağlayıcısına göndermeden önce
  abonelik sözleşmenizin buna izin verdiğini kontrol edin. Agent rapor metnini kopyalayıp yayımlamaz; yalnız kısa
  alıntılar kullanır.
- **Renk doğrulaması.** Pantone kodlarını fiziksel renk kartlarıyla, kumaşları tedarikçi numuneleriyle doğrulayın.
- **Taranmış PDF.** Görüntü olarak taranmış raporlardan metin çıkarılamaz; önce OCR uygulayın.
- **Örnek veri:** Trend notları ve brif kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk özette bilerek:
- alıntıyı rapordaki gibi yazmaz,
- raporda olmayan bir Pantone kodu (17-1463 TCX) uydurur,
- brifte olmayan "Ceket" grubunu ekler,
- silüet kategorisini atlar,
- kaçınılacaklar listesindeki payet ve metalik kumaşı önerir.

Kodun bunları yakalayıp düzelttirdiği doğrulanır.

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

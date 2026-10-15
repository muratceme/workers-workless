# Müşteri Yorumları (Fit Comments) Özeti · AI Agent

> Müşteri Temsilciliği (Merchandising) › Development Merchandiser · Workers / Workless

Müşterinin fit ve numune yorum e-postasını modelhane, kalıp, kumaş, baskı ve üretim ekipleri için **Türkçe, madde madde
aksiyon listesine** çevirir. Her madde e-postadaki satıra dayanır; ölçü değişiklikleri e-postada yazan değerlerle aktarılır,
inç değerleri kodla cm'ye çevrilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** E-posta satır satır numaralanır. Selamlama ve imza satırları ayrılır. Ölçüler (kesirli inç dahil:
   3/4", 1 1/2") bulunur.
2. **Özet (model):**
   - Model (style) başına genel karar: onay, yorumlu onay, ret
   - Aksiyonlar: bölge, yorumun Türkçesi, yapılacak iş, sorumlu birim, öncelik, ölçü değişikliği (POM, ölçülen, istenen,
     değişim)
   - Müşterinin istediği numune ve terminler
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Dayanak satırı olmayan ya da e-postada bulunmayan satıra dayanan madde
   - E-postada geçmeyen model no
   - Dayandığı satırda yazmayan ölçü değeri (model kendi çevirmiş ya da hesaplamışsa yakalanır)
   - Dayandığı satırda olmayan termin tarihi
   - Selamlama / imza dışında aksiyon listesine girmeyen satır
4. **Çevirme (kod):** İnç değişimler 2,54 ile çarpılarak cm karşılığı eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                      # örnek fit yorumu e-postası
python agent.py --girdi fit_comments.docx
python agent.py --girdi fit_comments.pdf --tur 3
```

E-postayı Outlook'tan .txt olarak kaydedebilir ya da metni bir .txt dosyasına yapıştırabilirsiniz.

## Çıktı

- `cikti/fit_aksiyonlari.xlsx`:
  - `Kararlar ve Terminler`
  - `Aksiyon Listesi`: birime göre sıralı; ölçü ve cm karşılığı, e-posta satırı ve "Durum" sütunu.
  - `E-posta`: her satırın nerede kullanıldığı; kullanılmayan satırlar kırmızı.
- `cikti/fit_aksiyonlari.md`: birim başlıklarıyla, ekiplere iletilecek liste.

## Dikkat

- **Taslaktır.** Modelhaneye iletmeden önce müşterinin orijinal e-postasıyla karşılaştırın. Ölçü değişikliklerini ölçü
  tablosuna (spec sheet) kendiniz işleyin; bu araç ölçü tablosunu güncellemez.
- **Ölçü yönü.** Değişimin yönünü (artır / azalt) model yorumlar; e-postadaki ifadeyle karşılaştırın.
- **Fotoğraf ve ekler okunmaz.** Yalnız e-posta metni işlenir; fotoğraflı yorumları ayrıca kontrol edin.
- **Veri gönderimi.** E-posta seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon ve e-posta adresleri maskelenir.
  Gerekirse yerel model (Ollama) kullanın.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek 3/4 inçlik değişimi kendisi cm'ye çevirir (1,9), bir yorum
satırını atlar ve termini yanlış yazar. Kodun bunları yakalayıp düzelttirdiği, kesirli inçleri doğru okuyup cm'ye çevirdiği
doğrulanır.

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

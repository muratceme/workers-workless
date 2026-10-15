# Konşimento Bilgi Çıkarma · AI Agent

> Deniz ve Hava Yolu Operasyon › Denizyolu Operasyon Uzmanı · Workers / Workless

Konşimento (Bill of Lading) belgesinden yükleyici, alıcı, ihbar adresi, gemi / sefer, limanlar, konteynerler ve yük
bilgilerini **yapılandırılmış kayda** (Excel + JSON) dönüştürür. Çıkarılan her bilgi kodla belgeyle karşılaştırılır;
konteyner numaraları ISO 6346 kontrol hanesiyle doğrulanır.

## Nasıl çalışır?

1. **Çıkarma (model):** Belgedeki alanlar, konteyner satırları ve TOPLAM satırı okunur.
2. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Metin alanı ya da sayı belgede geçmiyor (uydurma ya da yanlış okuma)
   - Konteyner numarası ISO 6346 kontrol hanesine uymuyor
   - Konteynerlerin kap / kg / CBM toplamı belgedeki TOPLAM satırıyla tutmuyor
   - Toplam konteyner adedi listeyle tutmuyor ya da bir konteyner iki kez yazılmış
3. **Uyarılar (kod):**
   - Emre yazılı (TO ORDER) konşimento
   - "Shipper's load, stow and count / said to contain" kaydı
   - Telex release / sea waybill ibaresi
   - Navlun ödeme şekli ya da orijinal adedi eksik
   - Düzenleme tarihi gemiye yükleme tarihinden önce
4. **Çeki listesi karşılaştırması (isteğe bağlı):** Konteyner bazında kap adedi ve brüt kg farkları, listede olup
   konşimentoda olmayan konteynerler.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek konşimento + çeki listesi
python agent.py --girdi konsimento.pdf
python agent.py --girdi konsimento.pdf --ceki ceki_listesi.xlsx --tur 3
```

## Çıktı

- `cikti/konsimento.xlsx`:
  - `Konşimento`: alanlar ve belgede var mı kontrolü, uyarılar.
  - `Konteynerler`: ISO 6346 sonucu, mühür, tip, kap, kg, CBM ve belge toplamı.
  - `Çeki Karşılaştırma`: yalnız çeki listesi verildiyse.
- `cikti/konsimento.json`: sisteme aktarılacak kayıt.

## Dikkat

- **Kontrol edin.** Gümrük beyanı, akreditif ibrazı ya da teslim işlemi öncesinde orijinal belgeyle karşılaştırın.
- **Taranmış PDF.** Metin içermeyen (taranmış) PDF'ler okunamaz; önce OCR uygulayın.
- **ISO 6346** yalnız numaranın biçimini ve kontrol hanesini doğrular; konteynerin gerçekten o sahibe ait olduğunu
  doğrulamaz.
- **Veri gönderimi.** Belge seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon ve e-posta maskelenir; firma adları
  ve adresler maskelenmez. Gerekirse yerel model (Ollama) kullanın.
- **Örnek veri:** Firmalar, gemi, numaralar ve konteynerler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek gemi adını belgede olmayan biçimde yazar ve bir konteynerin
kap adedini yanlış okur. Kodun bunları ve tutmayan toplamı yakalayıp düzelttirdiği doğrulanır.

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

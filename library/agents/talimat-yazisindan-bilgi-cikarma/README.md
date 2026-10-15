# Talimat Yazısından Bilgi Çıkarma · AI Agent

> Şube Bankacılığı › Gişe Yetkilisi · Workers / Workless

Mektup, faks veya e-postayla gelen **yazılı ödeme talimatlarından** gönderen, alıcı, IBAN, tutar (rakam ve yazı), para
birimi, işlem türü, açıklama ve imzalayanları çıkarır; her ödemeyi kontrol tablosuna aktarır. Çıktının ilk sayfası
**Müşteri Talimat Kontrolü** kod bloğunun girdi biçimindedir: imza yetkisi, hesap ve unvan kontrolü için oraya verin.

## Nasıl çalışır?

1. **Maskeleme (kod):** Belgedeki TR IBAN'lar `[IBAN-1]`, `[IBAN-2]` … yer tutucularıyla değiştirilir; **IBAN'lar
   sağlayıcıya gönderilmez.** Telefon, e-posta ve TCKN de maskelenir; `--gizle` ile verilen adlar `[KİŞİ-n]` olur.
2. **Çıkarma (model):** Her ödeme ayrı kayıt olur (bir mektupta birden fazla alıcı olabilir). Model IBAN alanına
   yalnız yer tutucuyu yazar ve hangisinin gönderen, hangisinin alıcı olduğunu metinden çıkarır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Unvan, açıklama, tarih, tutar, yazıyla tutar, referans veya imzalayan belgede aynen geçmiyor
   - IBAN alanında belgede olmayan bir değer var (uydurma IBAN kabul edilmez)
   - Gönderen ve alıcı IBAN'ı aynı; rakamla tutar boş veya sayı olarak okunamıyor
4. **Bulgular (kod, belgeyle ilgili):** Gerçek IBAN'lar geri konduktan sonra:
   - IBAN kontrol hanesi (ISO 13616 mod 97), TR IBAN uzunluğu ve rezerv alanı
   - Rakamla ve yazıyla tutar uyuşmuyor; yazıyla tutar yok veya okunamıyor
   - Alıcı IBAN'ı, imzalayan, işlem türü veya para birimi belgede yok
   - E-postayla gelmiş görünen talimat (ıslak imza yok)
   Durum: herhangi bir "Hata" → **Hatalı**; yalnız "Dikkat" → **Kontrol**; hiçbiri → **Uygun**.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                         # örnek iki talimat (mektup + e-posta)
python agent.py --girdi ./talimatlar                    # klasördeki tüm .pdf / .docx / .txt
python agent.py --girdi talimat.pdf --gizle "Ayşe Deneme" "Kemal Örnek"
```

Ardından imza ve hesap kontrolü için (Müşteri No sütununu doldurduktan sonra):

```bash
python main.py --girdi cikti/talimatlar.xlsx --hesaplar hesaplar.xlsx --yetkiler imza_yetkileri.xlsx   # musteri-talimat-kontrolu
```

## Çıktı

- `cikti/talimatlar.xlsx`:
  - `Talimatlar`: Talimat No, Tarih, Tür, Müşteri No (boş), Gönderen Unvan / IBAN, Alıcı Unvan / IBAN, Tutar, Tutar (Yazı),
    Para Birimi, Açıklama, İmzalayanlar + Durum, Bulgular, Kaynak Dosya. Mor hücreler yapay zekâ ile çıkarılmıştır.
  - `Bulgular`: talimat bazında bulgular ve modelin belirsizlik notları.
  - `Kontrol`: dosya bazında talimat sayısı, düzeltme turu, son denetim sonucu.
- `cikti/talimatlar.json`: aynı kayıtlar.

Talimat No: belgede sayı / referans varsa o (birden fazla ödemede `/1`, `/2` eklenir), yoksa `dosya adı-sıra`.

## Dikkat

- **İşlemden önce orijinal talimatla karşılaştırın.** Bu paket talimatı yalnız okur ve tabloya aktarır; imza yetkisi,
  hesap sahipliği, bakiye, limit ve müşteri kimliği kontrolü yapmaz. Bunlar için Müşteri Talimat Kontrolü kod bloğunu
  ve bankanızın prosedürlerini kullanın.
- **IBAN geçerliliği** yalnız biçim ve kontrol hanesini doğrular; hesabın var olduğunu veya alıcıya ait olduğunu
  doğrulamaz. Yalnız TR IBAN'lar maskelenir ve çıkarılır; yabancı IBAN'lar metinde açık kalır.
- **Taranmış PDF.** Metin içermeyen (taranmış) PDF'ler okunamaz; önce OCR uygulayın. El yazısı talimatlarda OCR
  hataları sık olur.
- **E-posta / faks talimatları** için bankanızın talimat sözleşmesi ve teyit prosedürü geçerlidir; paket bunu
  kontrol etmez, yalnız uyarır.
- **Veri gönderimi.** IBAN, telefon, e-posta ve TCKN dışında belge metni (unvanlar, kişi adları, tutarlar, açıklamalar)
  sağlayıcıya gönderilir. Kişi adlarını `--gizle` ile maskeleyin ya da yerel model (Ollama) kullanın.
- **Örnek veri:** Firmalar, kişiler, IBAN'lar ve tutarlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk çıkarımda bilerek alıcı unvanını belgedeki yazımdan farklı yazar ve bir IBAN
alanına yer tutucu yerine IBAN uydurur. Kodun bunları yakalayıp düzelttirdiği, IBAN'ların sağlayıcıya giden metinde hiç
bulunmadığı, rakam-yazı farkının ve hatalı kontrol hanesinin raporlandığı doğrulanır.

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

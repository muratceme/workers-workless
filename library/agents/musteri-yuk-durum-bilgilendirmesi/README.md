# Müşteri Yük Durum Bilgilendirmesi · AI Agent

> Karayolu Operasyon › Karayolu Operasyon Uzmanı · Gerisi Bende

Sefer durum tablosundan her müşteriye tek bir **yük durum e-postası** hazırlar: teslim edilen, zamanında ilerleyen ve geciken
yükler, son konum ve tahmini varış. E-postalar **gönderilmez**; Excel'de onay sütunuyla ve müşteri başına .txt dosyası
olarak hazırlanır.

## Nasıl çalışır?

1. **Hazırlık (kod):**
   - Her sefer için gecikme hesaplanır: ETA − planlanan teslim. `--tolerans` dakikayı aşarsa "Gecikmeli" (varsayılan 60).
   - Seferler müşteriye göre gruplanır.
   - Plaka, şoför ve müşteri e-posta adresi modele gönderilmez.
2. **Yazım (model):** Her müşteri için konu ve gövde yazılır. Geciken yüklerde verideki neden ve yeni tahmini varış
   kullanılır; neden yazılmamışsa uydurulmaz.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o e-posta modele geri gönderilir (varsayılan en fazla 2 tur):
   - Müşterinin bir sefer ya da yük numarası e-postada yok
   - Başka bir müşterinin numarası e-postaya karışmış
   - Tahmini varış saati yazılmamış
   - Veride olmayan sayı, saat ya da tarih eklenmiş
   - Gecikmesi olmayan müşteriye gecikme yazılmış
   - "Garanti" gibi taahhüt ifadesi kullanılmış

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                                     # örnek seferler
python agent.py --girdi seferler.xlsx --firma "Firma Adı" --iletisim "0212 000 00 00"
python agent.py --girdi seferler.xlsx --firma "Firma Adı" --iletisim "0850 000 00 00" --tolerans 30
```

Sefer durumunu TMS'nizden (taşıma yönetim sistemi) Excel olarak dışa aktarabilirsiniz. Tarih-saat sütunları
GG.AA.YYYY SS:DD biçiminde olmalıdır.

## Çıktı

- `cikti/yuk_durum.xlsx`:
  - `E-postalar`: alıcı, konu, gövde, geciken yük sayısı ve "Onay" sütunu.
  - `Sefer Durumu`: gecikme dakikası ve değerlendirme.
- `cikti/yuk_durum_epostalar/`: müşteri başına .txt e-posta taslağı (Kime, Konu, gövde).

## Dikkat

- **E-postalar gönderilmez.** Göndermeden önce sefer durumunu ve ETA'yı araç takip sisteminden teyit edin.
- **ETA tahminidir.** Tablodaki ETA'nın güncelliği sizin sisteminize bağlıdır; e-postalar taahhüt içermez.
- **Sıkı denetim.** Veride olmayan her sayı işaretlenir. Model tarihleri farklı yazarsa (ör. "15 Ekim") e-posta düzeltmeye
  gider; bu, uydurma saat ve tarihleri yakalamak içindir.
- **Veri gönderimi.** Müşteri adları ve sefer bilgileri seçtiğiniz yapay zekâ sağlayıcısına gönderilir; plaka, şoför ve e-posta
  adresi gönderilmez.
- **Örnek veri:** Firmalar, seferler ve plakalar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek gecikmesi olmayan müşteriye "gecikme olabilir" yazar, bir yük
numarasını unutur ve bir müşterinin e-postasına başka müşterinin sefer numarasını ekler. Kodun bunları yakaladığı,
gecikmeleri doğru hesapladığı ve plaka / şoför bilgisini modele göndermediği doğrulanır.

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

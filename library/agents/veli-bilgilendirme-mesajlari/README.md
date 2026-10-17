# Veli Bilgilendirme Mesajları · AI Agent

> Kayıt Kabul ve Öğrenci İşleri › Öğrenci İşleri Sorumlusu · Gerisi Bende

Duyuru, devamsızlık ve eksik ödev bilgilerinden **kişiselleştirilmiş veli mesajları** hazırlar. **Öğrenci ve veli bilgisi
yapay zekâ modeline gönderilmez.** Model yalnız mesaj türü ve duyuru metni için yer tutuculu şablon yazar; adları, sınıfı ve
sayıları kod doldurur. Mesaj göndermez; SMS sistemine aktarılacak CSV üretir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Listedeki mesaj türlerinden (Devamsızlık, Ödev, her duyuru) şablon kombinasyonları çıkarılır.
2. **Şablon (model):** Her kombinasyon için yer tutuculu kısa bir şablon yazılır: `{VELI}`, `{OGRENCI}`, `{SINIF}`,
   `{OZURSUZ}`, `{OZURLU}`, `{DERS}`, `{ODEV}`, `{OKUL}`, `{ILETISIM}`.
3. **Denetim (kod):** Aşağıdakilerden biri varsa şablonlar modele geri gönderilir (varsayılan en fazla 2 tur):
   - Zorunlu yer tutucu eksik ya da bilinmeyen yer tutucu var
   - Yer tutucuya ek getirilmiş ("{OGRENCI}'in"); ada göre yanlış ek çıkabilir
   - Duyurudaki tarih ya da saat şablonda aynen yok, ya da duyuruda olmayan sayı eklenmiş
   - Devamsızlık / ödev şablonunda sabit sayı yazılmış
   - Yargılayıcı ya da tehdit içeren ifade ("tembel", "başarısız", "sınıfta kalacak", "disiplin", "ceza")
   - En uzun değerlerle doldurulan mesaj `--azami` karakteri aşıyor (varsayılan 480)

   Kontrolden geçmeyen şablonla mesaj üretilmez.
4. **Doldurma (kod):** Telefonu olan veliler için mesajlar oluşturulur; telefonu olmayanlar "aranmalı" diye listelenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                          # örnek liste + duyurular
python agent.py --girdi liste.xlsx --duyurular duyurular.txt --okul "Okul Adı" --iletisim "0212 000 00 00"
python agent.py --girdi liste.xlsx --duyurular duyurular.txt --okul "…" --iletisim "…" --azami 300
```

Devamsızlık gün sayılarını e-Okul'dan aldığınız listeden doldurun. Duyuru dosyasında her duyuruyu "D1: …" biçiminde bir
satıra yazın ve listedeki "Duyuru" sütununa numarasını girin.

## Çıktı

- `cikti/veli_mesajlari.xlsx`:
  - `Şablonlar`: kontrol sonucu ve "Onay" sütunu.
  - `Mesajlar`: doldurulmuş mesajlar ve karakter sayısı.
  - `Gönderilemeyen`: telefonu olmayanlar ve kontrolden geçmeyen şablonlar.
- `cikti/veli_mesajlari_sms.csv`: Telefon; Mesaj.

## Dikkat

- **Şablonları okul idaresi onaylamalıdır.** Mesajlar gönderilmez; gönderimi okulunuzun SMS sistemi ya da e-Okul VBS üzerinden
  yaparsınız.
- **Devamsızlık bildirimleri.** Mevzuatın öngördüğü resmî veli bildirimleri (ör. ortaöğretimde özürsüz devamsızlıkla ilgili
  yazılı bildirim) e-Okul ve okul idaresi üzerinden yapılır; bu mesajlar bilgilendirme amaçlıdır, resmî bildirimin yerine
  geçmez.
- **Kişisel veri.** Öğrenci ve veli bilgisi modele gönderilmez; yalnız mesaj türleri ve duyuru metinleri gönderilir. CSV ve
  Excel çıktısını yalnız yetkili personel kullanmalıdır.
- **Örnek veri:** Öğrenciler, veliler ve telefonlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek devamsızlık şablonuna "10 günü aşarsa sınıfta kalacaktır" yazar,
ödev şablonunda yer tutucuya ek getirir ve iletişim telefonunu unutur, duyurudaki saati değiştirir. Kodun bunları yakalayıp
düzelttirdiği ve öğrenci / veli bilgisini modele göndermediği doğrulanır.

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

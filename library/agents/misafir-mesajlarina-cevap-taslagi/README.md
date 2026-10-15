# Misafir Mesajlarına Cevap Taslağı · AI Agent

> Ön Büro › Resepsiyonist · Workers / Workless

Misafirlerin e-posta ve mesajlarına, **otelin kendi bilgilerine dayanarak** ve misafirin dilinde cevap taslağı hazırlar.
Bilgi dosyasında cevabı olmayan soruları ve iptal, iade, şikâyet gibi insan onayı gereken mesajları ayrıca işaretler.
Taslaklar gönderilmez.

## Nasıl çalışır?

1. **Hazırlık (kod):** Otel bilgileri numaralanır (B1, B2 …). Misafir adı ve rezervasyon numarası mesajdan çıkarılır;
   telefon ve e-posta maskelenir.
2. **Taslak (model):** Her mesaj için misafirin dilinde cevap yazılır. Hitapta `{AD}` yer tutucusu kullanılır, kod gerçek
   adı doldurur. Kullanılan bilgi numaraları, cevaplanamayan sorular ve insan onayı gerekip gerekmediği belirtilir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o mesaj modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bilgi dosyasında olmayan bilgi numarası
   - Cevapta, kullanılan bilgilerde ve misafirin mesajında olmayan sayı, saat ya da ücret
   - Bilgiye dayanmayan "ücretsiz / garanti" ifadesi
   - "Onaylandı / confirmed" gibi kesin onay
   - İptal / iade / şikâyet içeren mesaj ya da cevaplanamayan soru insan onayına işaretlenmemiş
   - `{AD}` yer tutucusu yok

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek mesajlar + otel bilgileri
python agent.py --girdi mesajlar.xlsx --bilgi otel_bilgileri.txt
```

Otel bilgileri dosyasının her satırına bir bilgi yazın: giriş / çıkış saatleri, kahvaltı, transfer ücreti, otopark, evcil
hayvan, havuz, iptal kuralları, iletişim. Satırları "B1:", "B2:" diye numaralayabilirsiniz; numarasız satırlar sırayla
numaralanır.

## Çıktı

- `cikti/misafir_cevaplari.xlsx`:
  - `Cevap Taslakları`: insan onayı gerekenler en üstte; "Onay" sütunuyla.
  - `Otel Bilgileri`: her bilginin hangi mesajlarda kullanıldığı.

## Dikkat

- **Taslaklar gönderilmez.** Gönderilmeden önce okunup onaylanmalıdır. İnsan onayı gerekenler ilgili yetkiliye iletilmelidir.
- **Bilgi dosyasını güncel tutun.** Ücret ve saatler yalnız bu dosyadan alınır; dosya eskiyse cevaplar da eski olur.
- **Çeviri kontrolü.** Bilmediğiniz dillerdeki cevapları gönderiminden önce bir çeviri aracıyla kontrol etmeniz önerilir.
- **Veri gönderimi.** Mesaj metni ve otel bilgileri seçtiğiniz yapay zekâ sağlayıcısına gönderilir; misafir adı ve rezervasyon
  numarası çıkarılır, telefon ve e-posta maskelenir.
- **Örnek veri:** Otel bilgileri ve misafirler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek transfer ücretini yanlış yazar, iade edilemez rezervasyonun iade
talebini insan onayına işaretlemez ve geç çıkış talebini "onaylandı" diye cevaplar. Kodun bunları yakalayıp yalnız bu
mesajları yeniden yazdırdığı ve misafir adlarını modele göndermediği doğrulanır.

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

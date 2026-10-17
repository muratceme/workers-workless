# Ziyaret Notu Özeti · AI Agent

> Satış › Satış Temsilcisi · Gerisi Bende

Müşteri ziyareti sonrası alınan dağınık notları **CRM'e girilecek özete ve takip aksiyonlarına** çevirir. Her madde not
satırına dayanır; aksiyon terminleri nottaki tarih ifadelerinden kod tarafından doğrulanır.

## Nasıl çalışır?

1. **Hazırlık (kod):** Not dosyası ziyaretlere ayrılır (`Ziyaret: Müşteri | GG.AA.YYYY`) ve satırlar numaralanır.
   Telefon, e-posta ve kimlik numaraları maskelenir.
2. **Özet (model):** Her ziyaret için özet, ihtiyaçlar, itirazlar, rakip bilgisi, fırsat aşaması ve aksiyonlar
   (sorumlu: Biz / Müşteri, termin) yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o ziyaretin özeti modele geri gönderilir (varsayılan en fazla
   2 tur):
   - Maddenin dayanak satırı yok ya da geçersiz
   - Madde ya da tahmini tutarda notta olmayan bir sayı var (ör. 40.000 yerine 45.000)
   - Aksiyon termini notta yazmıyor ve ziyaret tarihine göre hesaplanamıyor. Kodun tanıdığı ifadeler: "yarın", "2 gün
     içinde", "3 iş günü", "haftaya", "cuma", "ay sonuna kadar", "gelecek ay", "5 Kasım"
   - "Rakip" ya da `--rakipler` listesindeki bir ad geçen satır rakip bilgisinde kullanılmamış
4. **Uyarılar (kod):** Termini olmayan aksiyonlar "termin belirsiz", bizim tarafımızdaki indirim / fiyat / garanti
   taahhüdü içeren aksiyonlar "yönetici onayı gerekir" diye işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek iki ziyaret
python agent.py --girdi notlar.txt --rakipler "X Ambalaj,Y Kutu"
```

Not dosyası örneği:

```
Ziyaret: Örnek Gıda A.Ş. | 13.10.2026
Aylık yaklaşık 40.000 adet karton koli ihtiyaçları var.
Numune talep ettiler; haftaya cuma gününe kadar göndereceğiz.
```

## Çıktı

- `cikti/ziyaret_ozetleri.xlsx`:
  - `CRM`: ziyaret başına özet, fırsat aşaması, tahmini tutar, ihtiyaçlar, itirazlar, rakip bilgisi, kontrol sonucu.
  - `Aksiyonlar`: iş, sorumlu, termin, uyarı ve dayandığı not satırları.
- `cikti/ziyaret_ozetleri_gorevler.csv`: CRM'e görev olarak aktarmak için Müşteri; Görev; Sorumlu; Termin.

## Dikkat

- **Kontrol edin.** Özetler CRM'e girmeden önce satış temsilcisi tarafından okunmalıdır. Fiyat ve indirim taahhütleri
  yönetici onayı olmadan müşteriye iletilmemelidir.
- **Kişisel veri.** Müşteri firma adı ve notlar modele gönderilir; telefon, e-posta ve kimlik numarası maskelenir.
  Görüşülen kişilerin adlarını notlara yazmamanız önerilir (unvan yeterlidir).
- **Örnek veri:** Firmalar ve görüşmeler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk özette bilerek:
- aylık ihtiyacı 45.000 adet yazar,
- rakip fiyat satırını atlar,
- fiyat teklifine notta olmayan bir termin (20.10.2026) verir.

Kodun bunları yakaladığı, yalnız sorunlu ziyareti yeniden yazdırdığı ve göreli tarih ifadelerini doğru hesapladığı
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

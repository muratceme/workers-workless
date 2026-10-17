# Prim Tahsilat Hatırlatma Mesajları · AI Agent

> Sigorta Tahsilat › Tahsilat Uzmanı · Gerisi Bende

Vadesi geçmiş prim taksitleri için **kademeli hatırlatma e-postası veya SMS taslakları** hazırlar. Mesajlar ya
sigortalıya (poliçe başına) ya da acenteye (acente başına, poliçe listesiyle) gider. Rakamları ve kademeyi kod belirler,
model yalnız metni yazar. Hukuki sonuç içeren cümleyi model yazamaz: onu siz kademeler dosyasına yazarsınız.

## Nasıl çalışır?

1. **Hazırlık (kod):**
   - Açık taksitler okunur: **Prim Alacak Yaşlandırma** kod bloğunun çıktısındaki "Açık Taksitler" sayfası ya da aynı
     sütunlu bir tablo.
   - Vadesi geçmiş taksitler poliçe veya acente bazında toplanır.
   - En uzun gecikmeye göre kademe seçilir. Örnek: 1-15 gün nazik, 16-45 gün ikinci, 46+ gün son hatırlatma.
   - Sigortalı adları sağlayıcıya gönderilmez; acente mesajlarında `[SİGORTALI-n]` olarak gider.
2. **Mesaj (model):** Kademenin tonuna uygun metin yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o mesaj modele geri gönderilir (varsayılan en fazla 2 tur):
   - Toplam tutar, her poliçe no ve her vade tarihi kodun biçimiyle mesajda yok; ya da veride olmayan bir sayı veya tarih
     var
   - Hukuki sonuç ya da baskı ifadesi var: iptal, fesih, teminatın durması, icra, dava, avukat, haciz, hukuki / yasal
     takip, faiz, gecikme zammı, ceza, kara liste, noter, ihtar…
   - Gövde hitapla başlıyor; sigortalı yer tutucusu kalmış; e-posta konusu boş
   - SMS sınırı aşılıyor. Hitap, zorunlu cümle, ödeme bilgisi ve imza çıkarıldıktan sonra kalan bütçe modele verilir;
     toplam sınır `--sms-sinir`, varsayılan 320 karakter.
4. **Metin (kod):** Mesaja sırasıyla eklenir:
   - "Sayın …" hitabı,
   - kademenin **zorunlu cümlesi** (yalnız sigortalı mesajlarında),
   - `--odeme` bilgisi,
   - imza.

## Kademeler dosyası

```
Kademe;En Az Gecikme (gün);Ton;Zorunlu Cümle
Nazik hatırlatma;1;Kibar ve bilgilendirici, ödemenin gözden kaçmış olabileceğini varsay;
İkinci hatırlatma;16;Kibar ama net, ödeme tarihini teyit etmesini iste;
Son hatırlatma;46;Resmî ve net, acil ödeme ya da iletişim iste;Ödenmeyen primler poliçenizin genel şartları ve ilgili mevzuat çerçevesinde teminatınızı etkileyebilir. …
```

**Zorunlu cümleyi hukuk / uyum biriminize onaylatın.** Örnekteki cümle genel bir ifadedir; branşa ve poliçe genel
şartlarına göre değişebilir. Bu agent prim ödenmemesinin sonuçları hakkında hukuki bilgi vermez.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek: 5 poliçe, sigortalıya e-posta
python agent.py --alici acente --kanal sms                              # acente başına SMS
python agent.py --girdi cikti/prim_alacak.xlsx --kademeler kademeler.csv --odeme "Online ödeme: ..." --imza "Kurgu Sigorta Tahsilat"
```

## Çıktı

- `cikti/hatirlatma_mesajlari.xlsx`:
  - `Mesajlar`: alıcı, kademe, en uzun gecikme, vadesi geçmiş toplam, poliçeler, konu, gönderime hazır metin, karakter
    sayısı, son kontrol.
  - `Kontrol`
- `cikti/hatirlatma_mesajlari.txt`: bütün mesajlar.

## Dikkat

- **Göndermeden önce** taksitlerin hâlâ açık olduğunu (rapor tarihinden sonraki tahsilatlar) kontrol edin.
- **Resmî ihtar değildir.** Noter ihtarnamesi için **İhtarname Taslağı** agent'ını ve hukuk birimini kullanın.
- **İletişim izni.** SMS ve e-posta gönderiminde ticari ileti ve KVKK kurallarına uyduğunuzdan emin olun; hatırlatma
  mesajları sözleşmeye dayalı bilgilendirme niteliğinde olmalıdır.
- **Örnek veri:** Poliçeler, sigortalılar ve acenteler kurgusaldır (Prim Alacak Yaşlandırma kod bloğunun örnek verisi).

## Testler

Testler gerçek API çağırmaz. Sahte model, 108 gün gecikmeli poliçenin ilk mesajına bilerek "teminatınız durabilir"
yazar. Kodun bunu yakalayıp yalnız o mesajı yeniden yazdırdığı doğrulanır. Ayrıca şunlar doğrulanır:
- kademe ve gruplama,
- eksik / yanlış tutar ve tarih yakalama,
- tehdit ifadesi (icra) yakalama,
- zorunlu cümlenin yalnız sigortalı mesajına eklenmesi,
- SMS bütçesi,
- sigortalı adlarının hiçbir mesajda geçmemesi.

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

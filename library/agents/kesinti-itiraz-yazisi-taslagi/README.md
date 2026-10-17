# Kesinti İtiraz Yazısı Taslağı · AI Agent

> Sağlık › Anlaşmalı Kurumlar Uzmanı · Gerisi Bende

Özel sağlık sigortası şirketlerinin veya kurumların hastane faturalarına yaptığı kesintileri tek tek inceler. Her
kesinti için itiraz önerisi verir (Evet / Hayır / Belge eksik) ve haklı olanlar için itiraz yazısı taslağı hazırlar.
Yazılar yalnız **sizin verdiğiniz belgelere ve anlaşma / SUT metnine** dayanır: yazıdaki her madde numarasını kod
dayanak metniyle karşılaştırır, uydurma madde kabul edilmez.

## Nasıl çalışır?

1. **Hazırlık (kod):** Okunan girdiler:
   - kesinti listesi (**Kesinti ve İtiraz Takibi** kod bloğunun girdi biçimi),
   - fatura numarasıyla başlayan destek belgeleri (`belgeler/F-201.txt`, `F-201_epikriz.pdf` …),
   - dayanak metni (anlaşma maddeleri, SUT maddeleri),
   - itiraz süreleri.
   Hasta adları (`--gizle`), TCKN, telefon, e-posta ve IBAN sağlayıcıya gönderilmez.
2. **İnceleme (model):** Her kesinti için öneri, gerekçe, dayanak maddeler, kullanılan belgeler ve eksik belgeler
   yazılır; "Evet" ise itiraz yazısının gövdesi de hazırlanır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o kesinti modele geri gönderilir (varsayılan en fazla 2 tur):
   - Yazıda dayanak metninde ya da belgelerde olmayan bir madde / SUT numarası / ek var
   - Kesinti kaydında, belgede ya da dayanakta olmayan bir sayı veya tarih var
   - "Evet" önerisi bir belgeye dayanmıyor; yazıda fatura no ya da kesinti tutarı aynen yok
   - "Belge eksik" önerisinde eksik belgeler yazılmamış; "Hayır" önerisinde yazı var
   - Kesinlik veya tehdit ifadesi var ("kesinlikle", "dava açılacaktır"); hasta yer tutucusu kalmış
4. **Süre (kod):** İtiraz süreleri verilmişse son itiraz tarihi = kesinti tarihi + süre. Süresi geçmiş ve 5 gün içinde
   dolacak itirazlar işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek 3 kesinti (2 itiraz, 1 belge eksik)
python agent.py --kesintiler kesintiler.xlsx --belgeler ./belgeler --dayanak anlasma.txt --sureler itiraz_sureleri.csv --gizle "Ayşe Deneme"
```

Dayanak metnine itirazda kullanmak istediğiniz anlaşma maddelerini ve ilgili mevzuat maddelerini kendi kaynağınızdan
(anlaşma metni, güncel SUT) kopyalayın. Agent mevzuatı kendisi bilmez ve aramaz.

Ardından takip için:

```bash
python main.py --kesintiler cikti/kesinti_itirazlari.xlsx      # Kesinti ve İtiraz Takibi kod bloğu klasöründe
```

## Çıktı

- `cikti/kesinti_itirazlari.xlsx`:
  - `Kesintiler`: girdi sütunları aynen (Kesinti ve İtiraz Takibi'ne geri verilebilir), ardından itiraz önerisi,
    gerekçe, dayanak maddeler, kullanılan belgeler, eksik belgeler, son itiraz tarihi, süre durumu, son kontrol.
    İtirazı gönderdiğinizde "İtiraz Tarihi" sütununu doldurun.
  - `İtiraz Yazıları`: önerisi "Evet" olan kesintilerin gönderime hazır yazıları.
  - `Kontrol`
- `cikti/kesinti_itirazlari.txt`: itiraz yazıları.

## Dikkat

- **Taslaktır.** Göndermeden önce faturalama ve ilgili tıbbi birimle kontrol edin. Tıbbi gereklilik değerlendirmesi
  hekime aittir; agent yalnız belgede yazan hekim gerekçesini aktarır.
- **Mevzuat güncelliği sizin sorumluluğunuzdadır.** Dayanak metnine koyduğunuz SUT veya anlaşma maddesinin yürürlükte
  olduğunu kontrol edin.
- **Sağlık verisi özel nitelikli kişisel veridir (KVKK md. 6).** Epikriz ve hekim notları modele gider. Hasta adlarını
  `--gizle` ile verin; mümkünse belgelerden kimlik bilgilerini çıkarın ya da yerel model (Ollama) kullanın.
- **Örnek veri:** Sigorta şirketleri, anlaşma maddeleri, fiyatlar ve faturalar kurgusaldır. Örnek anlaşma maddeleri
  gerçek bir anlaşmadan alınmamıştır.

## Testler

Testler gerçek API çağırmaz. Sahte model F-202 itirazında ilk turda bilerek dayanak metninde olmayan "Madde 4.5"e atıf
yapar. Kodun bunu yakaladığı ve yalnız o yazıyı yeniden yazdırdığı doğrulanır. Ayrıca şunlar doğrulanır:
- uydurma tutar, olmayan belge, eksik fatura no ve tehdit ifadesi yakalama,
- belgesiz "Evet" önerisinin reddi,
- itiraz süresi hesabı,
- hasta adının mesajlara gitmemesi.

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

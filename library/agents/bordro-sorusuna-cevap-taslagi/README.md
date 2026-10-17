# Bordro Sorusuna Cevap Taslağı · AI Agent

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Gerisi Bende

Çalışanların "maaşım neden düştü?", "fazla mesaim yatmamış", "bu kesinti ne?" gibi sorularına **bordro verisine
dayanan**, anlaşılır cevap taslakları hazırlar. Rakamları ve farkları kod hesaplar, model yalnız açıklar. Bordrodan
anlaşılamayan konular (puantaj, açıklaması olmayan kesinti) "insan kontrolü gerekir" diye işaretlenir.

## Nasıl çalışır?

1. **Hesap (kod):** Bordro dökümünden soruyu soran çalışanın sorulan ay ve önceki ay satırları bulunur. Her kalemin farkı
   hesaplanır ve olası nedenler işaretlenir:
   - **Neden:** vergi dilimi değişti; gün sayısı azaldı (eksik gün nedeniyle); asgari ücret GV istisnası değişti; fazla
     mesai / prim değişti.
   - **İnsan kontrolü:** açıklaması bordroda olmayan kesinti; soruda fazla mesai geçiyor ama bordroda yok (puantaj
     kontrolü); çalışan hata bildiriyor ama iki ay arasında fark yok.
2. **Cevap (model):** Kodun hesap tablosundaki rakamlarla soru sade bir dille cevaplanır. Model kendisi hesap yapmaz;
   ödeme veya düzeltme sözü vermez.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o cevap modele geri gönderilir (varsayılan en fazla 2 tur):
   - Cevaptaki bir sayı bordro satırlarında, kodun hesapladığı farklarda ya da soruda yok
   - "Bordroya eklenecektir", "iade edilecektir", "düzeltilecektir" gibi bir söz var
   - Kodun insan kontrolü işaretlediği bir konu insan kontrolüne gönderilmemiş, ya da kontrol nedeni yazılmamış
   - Cevap hitapla başlıyor, ad yer tutucusu kalmış ya da 1.500 karakteri geçiyor
4. **Gönderim metni (kod):** "Sayın Ad Soyad," hitabı ve `--imza` eklenir.

## Gizlilik

Sağlayıcıya yalnız soru metni ve iki aylık bordro rakamları gider. **Ad soyad ve sicil numarası gönderilmez.** Soru
metnindeki çalışan adı `[ÇALIŞAN]` yapılır; telefon, e-posta, TCKN ve IBAN da maskelenir. Maaş rakamları kişisel veridir:
şirket politikanız bulut sağlayıcısına izin vermiyorsa yerel model (Ollama) kullanın.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                           # örnek bordro ve 3 soru
python agent.py --bordro bordro_2026.xlsx --sorular sorular.txt --imza "Bordro Birimi"
python agent.py --bordro cikti/bordro.xlsx --sorular sorular.txt         # Brüt-Net Maaş Hesaplama çıktısı
```

**Bordro dökümü:** her satır bir çalışanın bir ayı.
- Gerekli sütunlar: `Ay` (Ekim / 10 / 2026-10) ve `Net Ücret`.
- Okunan diğer sütunlar: Sicil, Ad Soyad, Gün, Eksik Gün Nedeni, Brüt Ücret, Fazla Mesai, Prim, Toplam Brüt, SGK İşçi,
  İşsizlik İşçi, GV Matrahı, Kümülatif GV Matrahı, Vergi Dilimi %, Hesaplanan GV, AÜ GV İstisnası, Ödenecek GV,
  Ödenecek DV, Kesinti, Kesinti Açıklaması, Ele Geçen.
- **Brüt-Net Maaş Hesaplama** kod bloğunun kişi sayfaları da okunur (sayfa adı = çalışan).

**Sorular:**

```
Soru: 1001 | Ekim
Merhaba, ekim maaşım eylüle göre 1.750 TL az yattı. Bir hata mı var?

Soru: Zeynep Kurgu | Ekim
Bu ay maaşım neden 8 bin lira eksik geldi?
```

## Çıktı

- `cikti/bordro_cevaplari.xlsx`:
  - `Cevaplar`: soru, kategori, gönderime hazır taslak, insan kontrolü (Evet / Hayır) ve nedeni, kodun işaretledikleri,
    son kontrol.
  - `Hesap`: soru başına değişen kalemler (önceki ay, sorulan ay, fark).
  - `Kontrol`: düzeltme turu ve son denetim sonucu.
- `cikti/bordro_cevaplari.txt`: bütün taslaklar; insan kontrolü gerekenler başlıkta işaretli.

Bulunamayan çalışan, okunamayan ay ya da bordroda olmayan ay için cevap hazırlanmaz; neden Cevaplar sayfasına yazılır.

## Dikkat

- **Taslaktır.** Göndermeden önce bordro ve puantajla karşılaştırın; "İnsan Kontrolü: Evet" olanları kontrol etmeden
  göndermeyin.
- **Bordronun doğruluğunu kontrol etmez.** Model, bordroda yazan rakamları açıklar; vergi veya SGK hesabının doğru olup
  olmadığını denetlemez. Hesabı doğrulamak için **Brüt-Net Maaş Hesaplama** kod bloğunu kullanın.
- **Hukuki görüş değildir.** Mevzuat ve hak sorularında (kıdem, izin, fazla mesai ücreti oranı) uzman görüşü alın.
- **Örnek veri:** Çalışanlar kurgusaldır; örnek rakamlar Brüt-Net Maaş Hesaplama kod bloğunun 2026 parametreleriyle
  hesaplanmıştır.

## Testler

Testler gerçek API çağırmaz. Sahte model, fazla mesai sorusunun ilk cevabında bilerek "bir sonraki bordroya
eklenecektir" diye söz verir. Kodun bunu yakalayıp yalnız o cevabı yeniden yazdırdığı doğrulanır. Ayrıca şunlar
doğrulanır:
- kodun vergi dilimi, eksik gün, istisna, kesinti ve puantaj işaretleri,
- ad soyad ve sicilin hiçbir mesajda geçmemesi,
- uydurma sayı yakalama,
- Brüt-Net Maaş Hesaplama çıktısının okunması.

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

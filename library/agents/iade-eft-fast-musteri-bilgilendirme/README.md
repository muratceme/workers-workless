# İade EFT / FAST Müşteri Bilgilendirme · AI Agent

> Bankacılık Operasyon › Ödeme Sistemleri Operasyon Uzmanı · Gerisi Bende

Alıcı bankadan **iade gelen EFT / FAST ödemeleri** için gönderen müşteriye bilgilendirme mesajı taslağı hazırlar.
İade nedenini sade dille açıklar ve müşterinin ne yapması gerektiğini yazar. Ödeme listesini (tarih, tutar, alıcı, IBAN'ın
son 4 hanesi, referans) kod kurar. **Sağlayıcıya yalnız iade kodları ve banka açıklamaları gider**; müşteri no, alıcı
adı, IBAN ve tutar gönderilmez.

## Nasıl çalışır?

1. **Hazırlık (kod):** İade listesi (**İade Gelen EFT / FAST Analizi** kod bloğunun girdi biçimi) okunur, gönderen
   müşteri bazında toplanır. Farklı iade nedenleri (kod + açıklama) çıkarılır.
2. **Açıklama (model):** Her neden için sade açıklama ve müşterinin yapması gereken yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa model düzeltir (varsayılan en fazla 2 tur):
   - Bir neden açıklanmamış
   - Açıklamada sayı ya da yer tutucu var (tutar, tarih, IBAN ve referansı kod ekler)
   - Söz / kesinlik ifadesi var ("yatırılmıştır", "tekrar gönderilecektir", "masrafsız" …)
   - Müşterinin yapması gereken boş
4. **Mesaj (kod):** hitap, iade edilen ödemelerin listesi, neden açıklamaları, sizin `--bilgi` cümleniz (ör. iade
   tutarının hesaba aktarımıyla ilgili bankanızın standart ifadesi) ve imza.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek 4 iade, 3 müşteri
python agent.py --iadeler iadeler.xlsx --bilgi "İade tutarı hesabınıza aktarılmıştır." --imza "Örnek Bank Ödeme Sistemleri"
```

## Çıktı

- `cikti/iade_bilgilendirme.xlsx`:
  - `Mesajlar`: müşteri başına iade adedi, toplam tutar ve gönderime hazır mesaj.
  - `Neden Açıklamaları`: kod, banka açıklaması, adet, sade açıklama, müşterinin yapacağı (bir kez kontrol edip her
    müşteride kullanılır).
  - `Kontrol`
- `cikti/iade_bilgilendirme.txt`: bütün mesajlar.

## Dikkat

- **Göndermeden önce** iade tutarının müşteri hesabına aktarıldığını ve bankanızın müşteri iletişim kurallarını kontrol
  edin. Hesaba aktarım cümlesini model yazmaz; `--bilgi` ile siz verirsiniz.
- **Mesajda alıcı adı ve IBAN'ın son 4 hanesi** yer alır. Gönderim kanalınızın (SMS, e-posta, internet şubesi) kişisel
  veri kurallarına uygun olduğundan emin olun.
- **İade nedeni analizi** (kanal, şube, personel bazında) için **İade Gelen EFT / FAST Analizi** kod bloğunu kullanın.
- **Örnek veri:** Müşteriler, alıcılar ve IBAN'lar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda limit aşımı için "Ödemeniz yarın tekrar gönderilecektir" diye söz
verir. Kodun bunu yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- müşteri no, alıcı, IBAN, tutar ve referansın hiçbir mesaja gitmemesi,
- eksik / fazla neden, sayı ve yer tutucu denetimi,
- mesajın kodla kurulan ödeme listesi.

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

# BT Olay Sonrası Raporu · AI Agent

> Bilgi Teknolojileri › Sistem Uzmanı · Gerisi Bende

Olay sırasında tutulan zaman damgalı notlardan (sohbet dökümü, çağrı notları) **suçlamasız (blameless) olay sonrası
raporu** hazırlar. Rapor şu bölümlerden oluşur:
- özet ve etki,
- zaman çizelgesi,
- kök neden (Doğrulandı / Varsayım),
- katkıda bulunan etkenler,
- aksiyonlar ve açık sorular.

**Tespit, müdahale ve kesinti sürelerini kod hesaplar** ve SLA hedefiyle karşılaştırır.

## Nasıl çalışır?

1. **Hazırlık (kod):** Notlar okunur. Saatler başlıktaki tarihe bağlanır; gece yarısını geçen saatler ertesi güne
   yazılır. "Ad Soyad:" biçimindeki konuşanlar ve `--gizle` adları takma adla, telefon, e-posta, TCKN ve IBAN
   maskelenir. "İzleme:" gibi sistem kaynakları olduğu gibi kalır.
2. **Rapor (model):** Yalnız notlarda yazana dayanır; kişi yerine rol / ekip yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa model düzeltir (varsayılan en fazla 2 tur):
   - Zaman çizelgesinde notlarda olmayan damga; sırasız çizelge
   - Etki başlangıcı ve tespitin birer kez olmaması; hizmetin düzelmesinin tespitten ya da etkiden önce görünmesi
   - Notlarda olmayan sayı ya da tarih (süreleri model yazmaz)
   - Raporda kişi takma adı (suçlamasız kural)
   - Notta teyit yokken kök nedenin "Doğrulandı" sayılması
   - Notlardaki bir "Aksiyon:" satırının rapora alınmaması ya da notta olmayan aksiyon eklenmesi; sahibi boş aksiyon
4. **Süreler (kod):**
   - ilk belirti → tespit,
   - etki → tespit,
   - tespit → hizmet düzeldi,
   - kesinti (etki → hizmet düzeldi),
   - kalıcı çözüme kadar.

   `--sla-dk` verilirse kesinti süresi hedefle karşılaştırılır.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek olay (log diski doldu, SLA 60 dk)
python agent.py --notlar olay.txt --sla-dk 60 --gizle "Ayşe"
```

Not örneği:

```
Olay: INC-2026-0412 | Hizmet: ERP | Tarih: 14.10.2026
08:40 İzleme: erp-db01 log diskinde boş alan %5'in altına düştü uyarısı
09:05 Gülay Tekin: Muhasebeden ERP'de fatura kaydedilemiyor diye 3 çağrı geldi
09:26 Gülay Tekin: Kullanıcılar tekrar fatura kaydedebiliyor
Aksiyon: Disk uyarıları nöbet telefonuna da gitsin (Altyapı ekibi, 31.10.2026)
```

Saat satırında tarih de yazılabilir: `15.10.2026 01:10 Nöbetçi: …`

## Çıktı

- `cikti/olay_sonrasi_raporu.md`: rapor metni (wiki / Confluence / e-posta için).
- `cikti/olay_sonrasi_raporu.xlsx`:
  - `Zaman Çizelgesi`
  - `Aksiyonlar`: sahip ve termin; Durum sütunu "Açık" başlar, aksiyon takibine aktarılabilir.
  - `Süreler`
  - `Kontrol`

## Dikkat

- **Suçlamasız yaklaşım** kişileri korumak içindir; raporda kişi adı olmadığını yayımlamadan önce yine de kontrol edin.
  Mesaj içinde geçen tek ad (soyadsız) otomatik maskelenmez; `--gizle` ile verin.
- **Sistem bilgisi gönderilir.** Sunucu adı, IP, sürüm gibi bilgiler modele gider; gizliyse notlardan çıkarın ya da
  yerel model (Ollama) kullanın.
- **Süreler notlar kadar doğrudur.** Etki başlangıcı tam bilinmiyorsa en erken kanıt alınır ve açık soru olarak yazılır.
- **Örnek veri:** Olay, sunucular ve kişiler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda katkıda bulunan etkenlerde bir kişiyi takma adıyla suçlar. Kodun
bunu yakalayıp düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- konuşan maskeleme ve gece yarısı geçişi,
- damga, sıra, tür sayısı, sayı / tarih denetimleri,
- aksiyon kapsamı ve "Doğrulandı" kuralı,
- süre ve SLA hesabı,
- kişi adlarının mesajlara gitmemesi.

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

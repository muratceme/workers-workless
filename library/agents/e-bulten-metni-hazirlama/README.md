# E-bülten Metni Hazırlama · AI Agent

> Pazarlama › Dijital Pazarlama Uzmanı · Gerisi Bende

Kampanya bilgisinden **konu satırı alternatifleriyle e-bülten metni** hazırlar. Model içeriği yazar; kampanya bağlantısını,
gönderen bilgilerini ve abonelikten çıkma (ret) alt bilgisini **kod ekler**. E-posta göndermez; e-posta servisinize
yükleyeceğiniz HTML ve düz metin üretir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Kampanya bilgisi okunur. Gönderen unvanı, 16 haneli MERSİS no, iletişim bilgisi, ret bağlantısı,
   kampanya adı, bitiş tarihi ve kampanya bağlantısı zorunludur; eksikse agent çalışmaz.
2. **Metin (model):** Konu satırı alternatifleri (varsayılan 5), ön başlık, başlık, gövde, buton metni ve koşul özeti
   yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Konu satırı eksik ya da `--konu-azami` karakteri aşıyor (varsayılan 60)
   - Konu satırında tamamı büyük harf kelime (kampanya kodu hariç) ya da birden fazla ünlem var
   - Kampanya bilgisinde olmayan yüzde, tarih ya da tutar var
   - Bilgide "%30'a varan" yazarken metinde "%30 indirim" yazılmış (yanıltıcı indirim)
   - "Tüm ürünlerde", "ücretsiz", "hediye", "en ucuz", "garanti" gibi bilgide olmayan iddia var
   - Spam ifadesi ("bedava", "hemen kazan", "son şans") var ya da metne bağlantı (URL) yazılmış
   - Bitiş tarihi ya da kampanya kodu gövdede yok
4. **E-bülten (kod):** Buton kampanya bağlantısına bağlanır. Alt bilgiye gönderen unvanı, MERSİS no, iletişim bilgisi,
   ret açıklaması ve abonelikten çıkma bağlantısı eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                       # örnek sonbahar kampanyası
python agent.py --girdi kampanya.txt --konu-azami 50 --konu-sayisi 3
```

Kampanya dosyası için `ornek_veri/kampanya.txt` dosyasına bakın.

## Çıktı

- `cikti/e_bulten.html`: e-posta servisine (ör. Mailchimp, Brevo, kendi sisteminiz) yüklenecek HTML.
- `cikti/e_bulten.txt`: düz metin sürümü.
- `cikti/e_bulten.xlsx`: konu satırı alternatifleri (karakter sayısı ve "Seçim" sütunu), bölümler, kontrol sonucu.

## Dikkat

- **Ticari elektronik ileti kuralları.** 6563 sayılı Kanun ve Ticari İletişim ve Ticari Elektronik İletiler Hakkında
  Yönetmelik uyarınca ileti yalnız onayı olan alıcılara gönderilebilir. İletide gönderenin kimlik bilgisi (tacirlerde
  ticaret unvanı ve MERSİS numarası), erişilebilir bir iletişim bilgisi ve alıcının ret hakkını kullanabileceği bir
  bağlantı bulunmalıdır. Alıcı listenizin İleti Yönetim Sistemi'nde (İYS) onaylı olduğunu gönderimden önce kontrol edin.
  Yürürlükteki mevzuatı Ticaret Bakanlığı kaynaklarından teyit edin.
- **İndirim ve fiyat bilgisi.** Kampanya bilginizdeki oran ve koşulların doğruluğu sizin sorumluluğunuzdadır. İndirimli
  satış duyurularında reklam mevzuatına uyun.
- **Test gönderimi.** Göndermeden önce farklı e-posta istemcilerinde test e-postası alın.
- **Örnek veri:** Şirket, MERSİS numarası ve bağlantılar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek:
- "SONBAHAR FIRSATI: %30 indirim!!" konu satırı yazar (büyük harf, çift ünlem, "varan" düşürülmüş),
- "tüm ürünlerde" der,
- kampanya kodunu gövdeden çıkarır,
- koşul özetine bağlantı yazar.

Kodun bunları yakaladığı ve gönderen bilgilerinin modele gönderilmeden alt bilgiye eklendiği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6563 sayılı Elektronik Ticaretin Düzenlenmesi Hakkında Kanun | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6563.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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

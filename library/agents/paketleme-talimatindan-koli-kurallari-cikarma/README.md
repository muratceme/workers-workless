# Paketleme Talimatından Koli Kuralları Çıkarma · AI Agent

> Tekstil › Sevkiyat Sorumlusu · Gerisi Bende

Müşterinin paketleme talimatını (packing instruction, e-posta, el kitabı; Türkçe ya da İngilizce) okur:
- **koli parametreleri:** solid / asorti, asorti oranı, koli içi adet, koli ölçüsü, azami brüt ağırlık,
- **diğer zorunluluklar:** polybag, etiket, shipping mark, askı / katlama, belge ve süre, koli içeriği. Bunlar
  sevkiyat öncesi kontrol listesine dönüştürülür.

Koli ağırlığını sevk dosyasındaki adet ağırlıklarıyla **kod hesaplar**. Sınırı aşan model için koli içi adet önerir ve
**Koli Listesi (Packing List) Hazırlama** kod bloğunun komutunu hazırlar.

## Nasıl çalışır?

1. **Hazırlık (kod):** Talimat ve isteğe bağlı sevk dosyası okunur; telefon, e-posta, TCKN maskelenir.
2. **Çıkarma (model):** Parametreler ve gereklilikler talimattan kelimesi kelimesine kanıtlarıyla çıkarılır. Talimatta
   olmayan alışılmış değer yazılmaz.
3. **Denetim (kod):** Aşağıdakilerden biri varsa model düzeltir (varsayılan en fazla 2 tur):
   - Koli içi, asorti adetleri ya da azami brüt talimatta yazmıyor
   - Koli ölçüsü talimatta bu hâliyle yok (ör. talimat 60x40x30 derken 60x40x40)
   - Talimatta ölçü ya da kg sınırı varken alan boş
   - Asorti seti koli içine tam bölünmüyor; asorti bedeni talimatta yok; solid'de asorti oranı var
   - Kanıt talimatta aynen yok; bir parametrenin kanıtı eksik
   - "-malı / -meli / zorunlu / must / shall / should" içeren bir talimat cümlesi hiçbir yere alınmamış
4. **Hesap (kod):**
   - Brüt = koli içi × adet ağırlığı + dara.
   - Sınırı aşan modelde koli içi en fazla ⌊(azami − dara) / adet ağırlığı⌋, asortide set adedinin katına yuvarlanır.
   - Asorti bedeni sevk dosyasında yoksa uyarı verilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek talimat + örnek sevk
python agent.py --talimat talimat.pdf --sevk sevk.xlsx --dara 0,8
```

## Çıktı

`cikti/paketleme_kurallari.xlsx`:
- `Parametreler`: değer ve talimattaki kanıt, **kod bloğu komutu** (örnekte
  `python main.py --sevk sevk.csv --yontem asorti --asorti S=1,M=2,L=2,XL=1 --koli-ici 24 --koli-olcu 60x40x30 --dara 0.8`).
- `Gereklilikler`: kontrol listesi; "Kontrol Edildi" ve "Kontrol Eden" sütunları boştur.
- `Ağırlık Kontrolü`: model bazında koli brüt ağırlığı, sınır, durum, önerilen koli içi.
- `Açık Noktalar`: talimatta belirsiz kalanlar ve ağırlık uyuşmazlıkları (müşteriden onay alınacaklar).
- `Kontrol`

## Dikkat

- **Müşteri talimatı esastır.** Agent talimatı yorumlamaz; çelişen kuralları (ör. 24 adet koli içi ile 12 kg sınırı
  birlikte sağlanamıyorsa) açık nokta olarak verir. Müşteriden yazılı onay alın.
- **Dara** talimatta yoksa `--dara` değeri kullanılır; gerçek koli darasıyla çalıştırın.
- **Örnek veri:** Müşteri, talimat ve siparişler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda koli ölçüsünü talimattaki 60x40x30 yerine alışılmış 60x40x40 yazar.
Kodun bunu yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- sayı, asorti bölünebilirliği, boş alan, beden ve kapsam denetimleri,
- ağırlık hesabı ve set katına yuvarlanan öneri,
- kod bloğu komutu.

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

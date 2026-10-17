# İkinci El Araç İlan Metni · AI Agent

> Otomotiv › Satış › İkinci El Satış Danışmanı · Gerisi Bende

Araç kartı ve ekspertiz notlarından ikinci el ilan metni yazar:
- başlık,
- açıklama,
- donanım listesi,
- **ekspertiz özeti** (boyalı, lokal boyalı, değişen parçalar ve tramer).

**Ekspertizin eksiksiz ve doğru yansıtıldığını kod denetler.** Hasarı gizleyen ya da ekspertizle çelişen ilan
yayınlanmaz.

## Nasıl çalışır?

1. **Hazırlık (kod):** Araç kartları ve ekspertiz notları okunur. Boyalı / değişen parçalar ve tramer tutarı ayrıştırılır.
   Plaka ve şasi no gönderilmez; telefon, e-posta, TCKN maskelenir.
2. **Yazım (model):** Yalnız karttaki ve notlardaki bilgiyle dürüst, abartısız ilan.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o araç modele geri gönderilir (varsayılan en fazla 2 tur):
   - Ekspertiz özetinde eksik boyalı / lokal boyalı / değişen parça ya da tramer tutarı
   - Ekspertizle çelişen ifade: boya varken "boyasız"; değişen varken "değişensiz / değişen parça yoktur"; tramer ya da
     değişen varken "kazasız / tramersiz / hasar kaydı yoktur". "Hatasız, kusursuz, orijinal" her zaman uyarı verir.
   - Başlıkta model yılı, metinde km yok ya da farklı
   - Kartta ve notlarda olmayan sayı (km, fiyat, tramer, bakım km'si …)
   - Notlarda olmayan donanım
   - `--baslik-sinir` aşımı
4. **Ek satır (kod):** `--ek-satir` ile verilen sabit metin (yetki belgesi no, iletişim, adres) her ilanın sonuna eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 2 araç
python agent.py --araclar araclar.xlsx --ekspertiz ekspertiz.docx --arac Stok-002 --ek-satir "Yetki Belgesi No: …"
```

Ekspertiz notu örneği:

```
Araç: Stok-002
Boyalı: sol ön çamurluk
Lokal boyalı: yok
Değişen: ön kaput
Tramer: 45.000 TL
Bakım: Son bakım 145.000 km'de yapıldı, triger seti değişti.
Donanım: geri görüş kamerası, otomatik klima, hız sabitleyici
Not: Ön kaput değişimi 2024'teki hafif çarpma sonrası yetkili serviste yapıldı.
```

## Çıktı

- `cikti/ilan_metinleri.xlsx`: `İlanlar` (mor: model) ve `Kontrol`.
- `cikti/ilan_metinleri.txt`: araç başına yayına hazır metin (başlık, açıklama, donanım, ekspertiz, ek satır).

## Dikkat

- **İlan yayımlamadan önce** ekspertiz raporu ve tramer sorgusuyla karşılaştırın. Kod yalnız notlarınızdaki bilgiyle
  karşılaştırır; notlar eksikse ilan da eksik olur.
- **Mevzuat ve platform kuralları:** İkinci el araç ticareti ve ilanlarla ilgili yönetmelik ve ilan platformu kuralları
  değişebilir. İlanda bulunması gereken bilgileri güncel metinden kontrol edip `--ek-satir` ile ekleyin.
- **Örnek veri:** Araçlar, marka / model ve ekspertiz bilgileri kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda Stok-002'nin değişen ön kaputunu ekspertiz özetine yazmaz. Kodun
bunu yakalayıp düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- tramer tutarı, ekli yasak ifadeler ("değişen parça yoktur", "kazasızdır"),
- yıl / km / olmayan sayı / uydurma donanım / başlık sınırı,
- ek satır ve araç seçimi,
- notu olmayan araçta durma.

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

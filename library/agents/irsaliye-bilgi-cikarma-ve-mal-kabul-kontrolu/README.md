# İrsaliye Bilgi Çıkarma ve Mal Kabul Kontrolü · AI Agent

> Depo › Depo Sorumlusu · Gerisi Bende

PDF, taranıp metne çevrilmiş ya da e-postayla gelen **sevk irsaliyelerinden mal kabul kaydı** çıkarır ve açık
siparişlerle karşılaştırır. Tedarikçi ürünü kendi adı ve koduyla yazar ("İZOTONİK NaCl %0,9 500 ML"); model bunu
siparişinizdeki malzemeyle ("Serum fizyolojik %0,9 500 ml") eşleştirir. Fazla / kısmi teslim, siparişte olmayan ürün ve
son kullanma tarihi kontrollerini kod yapar. E-irsaliye XML'iniz varsa yapay zekâya gerek yoktur; bu agent XML'i
olmayan belgeler içindir.

## Nasıl çalışır?

1. **Hazırlık (kod):** İrsaliyeler ve açık sipariş listesi okunur. Telefon, e-posta, IBAN ve kimlik numaraları maskelenir.
2. **Çıkarım (model):** İrsaliye başlığı ve satırlar çıkarılır, her satır açık siparişteki malzemeyle eşleştirilir.
   - Başlık: no, tarih, sevk eden, alıcı, sipariş referansı.
   - Satır: stok kodu, ürün adı, miktar, birim, lot / seri, son kullanma tarihi (SKT).
   - Etken madde / ölçü / tip farklıysa eşleştirilmez.
3. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Metin, lot veya SKT belgede aynen geçmiyor; miktar belgede yazmıyor (ör. siparişteki miktarı yazmak)
   - Lot alanında bir tarih var: sütun kaymış, SKT olabilir
   - Malzeme kodu sipariş listesinde ya da irsaliyedeki siparişin altında yok
   - Aynı sipariş kalemi iki satıra verilmiş. Aynı ürünün farklı lotlarla iki satırda gelmesi hata sayılmaz.
4. **Mal kabul (kod):**

| Durum | Bulgu |
|---|---|
| Red / İnceleme | Açık siparişlerde yok; son kullanma tarihi geçmiş |
| Kontrol | Fazla teslim (kalan sipariş miktarından fazla; aynı kaleme gelen bütün irsaliyeler toplanır); SKT `--skt-ay` içinde doluyor (varsayılan 6 ay); SKT'li üründe lot yok |
| Kabul | Kısmi teslim (açık kalan miktarla birlikte bilgi olarak) ya da tam teslim |

"AA/YYYY" biçimindeki SKT ayın son günü sayılır. Kalan sipariş = sipariş miktarı − teslim alınan.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                        # örnek: ecza deposu ve ambalaj irsaliyesi
python agent.py --girdi ./irsaliyeler --siparis acik_siparisler.xlsx --skt-ay 12
python agent.py --girdi irsaliye.pdf --tarih 16.10.2026                # sipariş listesi olmadan yalnız çıkarım ve SKT
```

Ardından son kullanma takibi için (hastane / eczane):

```bash
python main.py --stok cikti/mal_kabul.xlsx      # Son Kullanma Tarihi Takibi kod bloğu klasöründe
```

SKT'si olmayan satırlar (ör. ambalaj) o kod bloğunda "okunamadı" uyarısıyla atlanır.

## Çıktı

`cikti/mal_kabul.xlsx`:
- `Teslim Alınanlar`: her irsaliye satırı.
  - İlk sütunlar **Son Kullanma Tarihi Takibi** girdi sütunlarıyla aynı: Malzeme Kodu, Malzeme Adı, Lot No, Son Kullanma
    Tarihi, Miktar.
  - Diğer sütunlar: birim, sipariş no, sipariş kalanı, fark, irsaliye no, tedarikçinin kodu ve ürün adı, durum, bulgular.
- `İrsaliyeler`: irsaliye başına başlık bilgileri, kabul / kontrol / red sayıları, model notları ve son denetim sonucu.
- `Bilgi`

## Dikkat

- **Fiziksel kontrolün yerine geçmez.** Malı sayım, ambalaj ve etiket kontrolü yapmadan kabul etmeyin. Belge ile gelen mal
  farklı olabilir.
- **İlaç ve tıbbi malzemede** eşleşmeyi (etken madde, doz, form, ambalaj) eczacı kontrol etmelidir. Model ticari adı
  etken maddeyle eşleştirebilir ama farklı form veya dozu her zaman ayırt edemeyebilir.
- **SKT kabul kriteri** kurumunuza göre değişir (ör. en az 12 ay); `--skt-ay` ile ayarlayın.
- **Örnek veri:** Firmalar, ürün kodları, lotlar ve siparişler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ecza deposu irsaliyesinin ilk çıkarımında bilerek:
- kayan sütun yüzünden SKT'yi lot olarak okur,
- miktarı siparişten alır (650 yerine 600),
- seftriaksonu parasetamol kalemine eşleştirir,
- olmayan bir kod uydurur.

Kodun dört sorunu da yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- fazla / kısmi teslim; birden çok irsaliyenin aynı kaleme toplanması,
- SKT yakın / geçmiş kuralları,
- çıktının Son Kullanma Tarihi Takibi sütunlarıyla başlaması.

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

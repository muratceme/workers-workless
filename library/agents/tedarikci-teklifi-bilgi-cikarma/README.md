# Tedarikçi Teklifi Bilgi Çıkarma · AI Agent

> Satın Alma › Satın Alma Uzman Yardımcısı · Workers / Workless

Tedarikçilerden e-posta, PDF veya Word ile gelen fiyat tekliflerini **karşılaştırma tablosuna** aktarır. Her teklif
kalemini talep listenizdeki malzemeyle eşleştirir: tedarikçi "A42 V kayış", talep "V kayışı A-42" yazsa da eşleşir.
Ödeme, teslim, garanti ve geçerlilik ifadelerini gün ve aya kod çevirir. Çıktı doğrudan **Teklif Karşılaştırma** kod
bloğuna verilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Teklif belgeleri ve talep listesi okunur. Telefon, e-posta, IBAN ve kimlik numaraları maskelenir.
2. **Çıkarım (model):** Her teklif için şunlar çıkarılır:
   - tedarikçi, teklif no ve tarihi;
   - kalemler ve talep listesindeki karşılıkları (muadil / marka farkı not edilir; ölçü veya tip farkı varsa eşleştirilmez);
   - para birimi ve KDV dahil / hariç bilgisi;
   - ödeme, teslim, nakliye, garanti ve geçerlilik ifadeleri, belgedeki yazımıyla.
3. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Metin ya da ifade belgede aynen geçmiyor
   - Bir sayı belgede yazmıyor: hesaplanmış ya da uydurulmuş olabilir
   - Malzeme kodu talep listesinde yok ya da aynı kod iki kaleme verilmiş
4. **Sayılar ve bulgular (kod):** Gün, ay ve tarihleri model değil kod hesaplar:

| İfade | Kodun hesabı |
|---|---|
| "Peşin", "nakit", "sipariş ile birlikte" | Vade 0 gün |
| "faturadan itibaren 60 gün", "2 ay vade" | 60 gün |
| "5 iş günü" | Bugünden 5 iş günü sonrasına kadar geçen takvim günü (cumartesi-pazar atlanır, resmî tatil hesaba katılmaz) |
| "Teslim tarihi: 24.10.2026" | Bugünden o tarihe kadar gün |
| "2-3 hafta" | Üst sınır: 21 gün |
| "2 yıl garanti" | 24 ay |
| "Teklif 15 gün geçerlidir" | Teklif tarihi + 15 gün |
| "nakliye bize aittir", "fabrikanıza teslim" | Nakliye 0 |

| Önem | Bulgu |
|---|---|
| Hata | Kalem tutarı ≠ miktar × birim fiyat; teklifin geçerlilik süresi dolmuş; teslim tarihi geçmiş |
| Dikkat | Talep miktarından farklı miktar; KDV dahil fiyat (karşılaştırmadan önce KDV hariçe çevirin) ya da KDV belirtilmemiş |
| Dikkat | Nakliye bedeli belirsiz ("alıcıya aittir"); ödeme vadesi ya da teslim süresi yok |
| Dikkat | 3 gün içinde geçerliliği bitecek teklif; talep listesiyle eşleşmeyen kalem |
| Bilgi | Teklif verilmeyen kalemler, döviz teklif, garantisi yazmayan teklif, iş gününden çevrilen termin |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek talep (4 kalem) ve 3 teklif
python agent.py --girdi ./teklifler --talep talep.xlsx
python agent.py --girdi ./teklifler --talep talep.xlsx --tarih 20.10.2026     # termin ve geçerlilik bu tarihe göre
```

Talep listesi örneği (`.csv` noktalı virgülle ya da `.xlsx`):

```
Malzeme Kodu;Malzeme;Miktar;Birim
K-01;Rulman 6205-2RS;200;adet
```

Ardından karşılaştırma için:

```bash
python main.py --girdi cikti/teklifler.xlsx --kur USD=41,20 --faiz 40     # Teklif Karşılaştırma kod bloğu klasöründe
```

## Çıktı

`cikti/teklifler.xlsx`:
- `Teklif Listesi`: tedarikçi × kalem satırları, **Teklif Karşılaştırma** girdi sütunlarıyla.
  - Malzeme sütununda talep listesindeki ad yazılır; tedarikçinin yazdığı ad Not sütunundadır.
  - Nakliye tedarikçinin ilk satırına yazılır.
  - **Kalite Puanı** ve **Teknik Uygun (E/H)** sütunlarını siz doldurun. Boş Teknik Uygun "uygun" sayılır; geçerliliği
    dolmuş veya şartnameye uymayan teklifler için H yazın.
- `Teklifler`: teklif başına belgedeki ödeme / teslim / nakliye / garanti / geçerlilik ifadeleri, kodun hesapladığı
  değerler, özel şartlar, durum (Hatalı / Kontrol / Uygun) ve bulgular.
- `Kontrol`: dosya başına düzeltme turu ve son denetim sonucu.

Ayrıca aynı bilgiler `cikti/teklifler.json` dosyasına yazılır.

## Dikkat

- **Eşleşmeleri kontrol edin.** Muadil ürün, farklı marka, farklı boşluk sınıfı (ör. C3) veya farklı kaplama teknik olarak
  uygun olmayabilir. Model eşleştirme gerekçesini Not sütununa yazar; teknik uygunluk kararı sizindir.
- **Bu bir satın alma kararı değildir.** Fiyatlar teklifin aslıyla karşılaştırılmalı, şartname uyumu, numune ve
  referanslar ayrıca değerlendirilmelidir.
- **İş günü hesabı** resmî ve dinî bayramları bilmez; uzun tatil dönemlerinde termin daha uzun olabilir.
- **Kişisel veri.** Firma adları, kalemler ve fiyatlar modele gönderilir. Ticari sır niteliğindeki tekliflerde yerel model
  (Ollama) kullanmayı düşünün.
- **Örnek veri:** Firmalar, teklifler ve fiyatlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model, Örnek Endüstriyel teklifinin ilk çıkarımında bilerek:
- bir malzeme adını talep listesindeki yazımla "düzeltir",
- olmayan bir kod (K-9) uydurur,
- aynı kodu iki kaleme verir,
- belgede olmayan nakliye tutarı (0) yazar.

Kodun bunları yakalayıp modele geri gönderdiği doğrulanır. Ayrıca şunlar yakalanır:
- Kurgu Teknik'te tutar hatası (100 × 3,25 = 325, teklifte 352),
- Deneme Hırdavat'ta süresi dolmuş geçerlilik ve eksik miktar.

Peşin / iş günü / tarih / aralık ifadelerinin çevrimi ve çıktının Teklif Karşılaştırma sütunlarıyla aynı olduğu da
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

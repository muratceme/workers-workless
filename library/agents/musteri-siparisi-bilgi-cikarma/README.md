# Müşteri Siparişi Bilgi Çıkarma · AI Agent

> Üretim Planlama › Üretim Planlama Uzmanı · Workers / Workless

Müşterilerden e-posta veya PDF ile gelen siparişleri (PO) **planlama tablolarına** aktarır. Satırlar şunları içerir:
ürün, renk, beden dağılımı, miktar, fiyat ve teslim tarihi. Her satır ürün listenizle eşleştirilir: müşterinin kodu veya
açıklaması sizin ürün kodunuza bağlanır. Toplam ve termin kontrollerini kod yapar. Çıktı sitedeki üç planlama kod
bloğuna doğrudan girer.

## Nasıl çalışır?

1. **Hazırlık (kod):** Sipariş belgeleri ve ürün listesi okunur. Telefon, e-posta, IBAN ve kimlik numaraları maskelenir.
2. **Çıkarım (model):**
   - müşteri, PO no ve tarihi, para birimi, teslim şekli;
   - her ürün × renk satırı: müşteri kodu, açıklama, eşleşen ürün kodu, beden dağılımı, miktar, birim fiyat, teslim
     tarihi;
   - belgedeki genel toplam.
   Model toplamları "düzeltmez", belgede yazdığı gibi aktarır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Metin, renk veya beden belgede aynen geçmiyor (ör. "Navy" yerine "Lacivert")
   - Sayı belgede yazmıyor (ör. tutarsız toplamı düzeltip 1.500 yazmak)
   - Teslim tarihi okunmuyor ya da belgede yok
   - Ürün kodu listede yok
4. **Bulgular (kod):**

| Önem | Bulgu |
|---|---|
| Hata | Beden toplamı ≠ satır toplamı; satırlar toplamı ≠ belgedeki genel toplam; teslim tarihi geçmiş |
| Dikkat | Teslime `--min-gun`den (varsayılan 21) az var; ürün listesinde karşılığı yok (yeni ürün); teslim tarihi yok; fiyat var ama para birimi yok |
| Bilgi | Birim fiyat siparişte yok |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek: tekstil PO ve metal parça siparişi
python agent.py --girdi ./siparisler --urunler urun_listesi.xlsx --min-gun 30
```

## Çıktı

`cikti/musteri_siparisleri.xlsx`:

| Sayfa | İçerik | Verilebileceği kod bloğu |
|---|---|---|
| Sipariş Satırları | Bütün bilgiler, durum (Hatalı / Kontrol / Uygun), bulgular | — |
| Üretim Siparişleri | Sipariş No, Müşteri, Ürün, Miktar, Termin, Öncelik | Haftalık Üretim Çizelgesi (`--siparisler`) |
| Beden Dağılımı | Sipariş No, Model, Renk, bedenler, Toplam | Kumaş ve Aksesuar İhtiyaç Hesabı (`--siparis`) |
| T&A Siparişleri | Sipariş No, Müşteri, Model, Adet, PO Tarihi, Sevk Tarihi | Sipariş Termin Planı (T&A) |
| Kontrol | Belge başına genel toplam, sipariş bulguları, model notları, son denetim | — |

Notlar:
- Kod blokları ilk sayfayı okuduğu için ilgili sayfayı ayrı bir dosyaya kopyalayarak verin.
- Üretim Siparişleri'ndeki sipariş numaraları satır başına tekildir (SP-88213/1, /2 …).
- Ürün adı ürün listenizdeki addır; eşleşmeyen satırda müşterinin açıklaması yazılır.

## Dikkat

- **Siparişi onaylamadan önce** PO'nun aslıyla karşılaştırın. Beden tablolarında sütun kayması olabilir; kod beden
  toplamını kontrol eder ama hangi bedenin kaydığını bilemez.
- **Eşleşme kontrolü.** Müşteri kodu sizin kodunuzla aynı görünse de çizim revizyonu, renk kodu veya kumaş farklı
  olabilir; model bu farkları notlara yazar.
- **Termin teyidi.** Kodun kısa termin uyarısı kapasite hesabı değildir; Haftalık Üretim Çizelgesi veya Kapasite /
  Darboğaz Analizi ile kontrol edin.
- **Örnek veri:** Müşteriler, siparişler ve ürünler kurgusaldır. Örnek PO'da White satırının beden toplamı (1.500)
  satır toplamını (1.600) bilerek tutmaz.

## Testler

Testler gerçek API çağırmaz. Sahte model Kurgu Retail PO'sunun ilk çıkarımında bilerek:
- tutarsız satır toplamını 1.500'e "düzeltir",
- "Navy"yi "Lacivert" diye çevirir,
- olmayan bir ürün kodu (SW-200) yazar.

Kodun üçünü de yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- belgenin kendi tutarsızlığının bulgu olarak raporlanması,
- genel toplam, kısa / geçmiş termin, yeni ürün ve fiyat yok kuralları,
- üç planlama sayfasının kod blokları sütunlarıyla aynı olması.

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

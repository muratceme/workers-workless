# Stok Sayım Fark Analizi · Kod Bloğu

> Lojistik ve Depo › Depo Sorumlusu · Workers / Workless

Sayım sonuçlarını sistem (ERP) stoğuyla karşılaştırır. Farkları lokasyon ve ürün bazında miktar ve tutar olarak
raporlar, olası nedenleri işaretler ve ikinci sayım listesini çıkarır. İnternete bağlanmaz.

## Ne yapar?

**Fark hesabı**
- **Satır bazında fark:** Her lokasyon ve stok kodu için `sayılan − sistem` hesaplanır.
- **Sayım kartlarının birleştirilmesi:** Aynı lokasyon ve kod için birden fazla sayım satırı (farklı ekip ya da
  kart) toplanır.
- **Bekleyen hareketler:** Sayım anında sisteme işlenmemiş hareketler (sevk edilmiş ama irsaliyesi işlenmemiş mal,
  üretim fişi vb.) verilirse, sistem stoğu önce bu hareketlerle düzeltilir.
- **Ürün bazında mahsup:** Bir lokasyondaki eksik, aynı ürünün başka lokasyondaki fazlasıyla kapanıyorsa
  "yer karışıklığı" olarak işaretlenir. Muhasebe tutarı mahsup sonrası net farktan hesaplanır.

**Olası neden ipuçları (kural tabanlı)**

Bunlar kesin tespit değildir, yerinde kontrol önerisidir.

| İpucu | Kural |
|---|---|
| Hiç sayılmamış | Sistemde stok var, sayım 0 |
| Fazla/eksik sıfır | 30 ↔ 300 gibi 10/100/1000 kat |
| Rakam yer değiştirmesi | 54 ↔ 45 (aynı rakamlar, fark 9'un katı) |
| Koli/adet karışıklığı | Biri diğerinin koli içi adet katı; ya da fark tam koli |
| Varyant karışıklığı | Benzer kodlarda (BJ-1000-KIR / BJ-1000-MAV) aynı lokasyonda eşit ve ters fark |
| Negatif sistem stoğu | Çıkış, girişten önce kaydedilmiş |
| Kartı olmayan kalem | Yanlış okunmuş kod veya açılmamış kart |

**Değerlendirme**
- **Tolerans:** `--tolerans` ile (ör. %0,5 fire) kabul edilebilir farklar ayrılır.
- **Envanter doğruluğu:** Farkı tolerans içinde kalan satırların tüm satırlara oranıdır.
- **İkinci sayım listesi:** Tolerans dışında kalan ve şu koşullardan birini taşıyan satırlar listeye alınır:
  - tutarı `--esik` (varsayılan 5.000 TL) veya üzeri,
  - oranı `--esik-oran` (varsayılan %10) veya üzeri,
  - hiç sayılmamış,
  - stok kartı yok.

  Liste, kör sayım için boş sütunlarla gelir.
- **Muhasebe özeti:** 197 Sayım ve Tesellüm Noksanları ve 397 Sayım ve Tesellüm Fazlaları için tutar önerisi
  verir. Kesin kayıt için mali müşavirinize danışın.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                     # örnek verilerle
python main.py --sayim sayim.xlsx --sistem stok.xlsx
python main.py --sayim sayim.xlsx --sistem stok.xlsx --bekleyen bekleyen.xlsx --tolerans 0.5 --esik 10000
```

**Girdi dosyaları**
- **Sistem stoğu:** ERP'nin "lokasyon bazında stok durumu" raporu. Gerekli sütunlar Lokasyon, Stok Kodu, Stok Adı,
  Sistem Miktarı ve Birim Maliyet. Birim, Kategori ve Koli İçi isteğe bağlıdır.
- **Sayım:** El terminali veya sayım kartı dökümü. Gerekli sütunlar Lokasyon, Stok Kodu ve Sayılan Miktar.
- **Bekleyen hareketler:** Gerekli sütunlar Stok Kodu ve Miktar. Lokasyon isteğe bağlıdır. Miktar işaretli
  verilebilir (+ giriş, − çıkış) ya da Yön sütunuyla (Giriş/Çıkış) belirtilebilir.
- **Lokasyonsuz dosyalar:** Dosyalardan birinde lokasyon yoksa karşılaştırma stok kodu bazında yapılır.
- **Başlıklar:** Büyük/küçük harf ve Türkçe karakter farkı önemsizdir. "SAYILAN MİKTAR" ile "Sayilan Miktar"
  aynı başlık sayılır.

## Çıktı

`Özet` · `Farklar` · `Ürün Bazında` · `İkinci Sayım` · `Muhasebe` · `Bilgi`

## Dikkat

- **Sayım anı:** Sayım sırasında mal giriş-çıkışı durdurulmalıdır. Durdurulamıyorsa, sayım saatindeki işlenmemiş
  hareketler bekleyen dosyasına yazılmalıdır.
- **Belgesiz eksikler:** Nedeni belgelenemeyen eksikler vergi incelemesinde sorun yaratabilir. Fire, zayi ve
  hırsızlık tutanakla belgelenmelidir. KDV ve gider yönünden mali müşavirinize danışın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

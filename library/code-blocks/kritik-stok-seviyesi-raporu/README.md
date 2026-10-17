# Kritik Stok Seviyesi Raporu · Kod Bloğu

> Eczane ve Tıbbi Malzeme › Hastane Eczacısı · Gerisi Bende

İlaç ve tıbbi malzemelerin **günlük tüketimi ve tedarik süresinden kritik stok seviyesini** hesaplar; stoksuz, kritik
seviyenin altında ve emniyet stoğunun altındaki kalemleri listeler, **sipariş önerisi** çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Günlük tüketim | Günlük Tüketim sütunu · ya da Aylık Tüketim ÷ 30 · ya da `--tuketim` dosyasındaki dönem tüketimi ÷ `--gun` |
| Emniyet stoğu | Günlük tüketim × emniyet günü (satırda yoksa `--emniyet`, varsayılan 7) |
| Kritik seviye | Günlük tüketim × (tedarik süresi + emniyet günü) — sipariş verilmesi gereken seviye |
| Stok günü / tahmini tükenme | Mevcut stok ÷ günlük tüketim; rapor tarihi + stok günü |
| Önerilen sipariş | Günlük tüketim × (tedarik süresi + emniyet günü + `--kapsam`, varsayılan 30) − mevcut − bekleyen sipariş; yukarı yuvarlanır, yalnız Stoksuz ve Kritik altında kalemler için |

| Durum | Kural |
|---|---|
| Stoksuz | Mevcut stok sıfır |
| Kritik altında | Mevcut + bekleyen sipariş < kritik seviye |
| Sipariş yolda | Mevcut < kritik seviye ≤ mevcut + bekleyen |
| Yeterli | Mevcut ≥ kritik seviye |

**Uyarılar:** stoksuz ve emniyet stoğunun altındaki kalemler (Yüksek) · kritik seviyenin altındakiler (hayati
işaretliyse Yüksek, değilse Orta) · siparişi yolda olup tedarik süresinden önce tükenecekler · tüketimi ya da tedarik
süresi bilinmeyenler.

**Örnek:** Ondansetron aylık 30 ampul → günde 1; tedarik 7 + emniyet 7 gün → kritik seviye **14**. Mevcut 12 → kritik
altında; öneri 1 × (7 + 7 + 30) − 12 = **32 ampul**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                   # örnek veri (30.09.2026)
python main.py --stok eczane_stok_durumu.xlsx --emniyet 10 --kapsam 45
python main.py --stok eczane_stok_durumu.xlsx --tuketim son_3_ay_tuketim.xlsx --gun 90 --tedarik 7
```

Zorunlu sütunlar: Malzeme Kodu ve Mevcut Stok. Tedarik süresi satırda yoksa `--tedarik` kullanılır; ikisi de yoksa kalem
"Hesaplanamadı" olur. Hayati sütununa Evet yazılan kalemler önce sıralanır ve daha yüksek önemle uyarılır.

## Çıktı

`cikti/kritik_stok_raporu.xlsx`:

- `Kritik Stok`: stoksuz, kritik altında ve siparişi yolda kalemler; durum, kritik seviye, stok günü, tükenme tarihi,
  önerilen sipariş ve tutarı.
- `Tüm Malzemeler`: aynı göstergeler, tüm kalemler.
- `Uyarılar`.

## Dikkat

- **Kural tabanlı bir hesaptır.** Emniyet günü sabit bir süre olarak alınır; tüketimdeki dalgalanmaya göre istatistiksel
  emniyet stoğu hesaplanmaz. Hayati ilaçlar, acil ve yoğun bakım stokları için kurumun belirlediği asgari stok
  düzeylerini esas alın.
- **Önerilen sipariş taslaktır.** Ambalaj miktarı, minimum sipariş, ihale / sözleşme koşulları, bütçe ve depo kapasitesi
  gözetilmez; satın alma kararı eczacı ve satın alma biriminindir.
- **Tüketim verisi güncel olmalı.** Salgın, mevsim, yeni protokol ya da servis açılışı tüketimi değiştirir. Piyasada
  bulunamayan ilaçlarda tedarik süresi gerçeği yansıtmayabilir.
- **Depo bazında takip** gerekiyorsa her depo için ayrı dosyayla çalıştırın; aynı malzeme ikinci kez yazılırsa ikinci
  satır alınmaz.
- **Örnek veri:** Malzeme kodları, stoklar ve fiyatlar kurgusaldır.

## Testler

- Emniyet stoğu, kritik seviye, stok günü, tükenme tarihi, önerilen sipariş ve tutarı.
- Stoksuz / kritik altında / sipariş yolda / yeterli ayrımı, hayati kalem önceliği, varsayılan tedarik süresi, dış
  tüketim dosyası, okuma ve komut satırı.

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

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.

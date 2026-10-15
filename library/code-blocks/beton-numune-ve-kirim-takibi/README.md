# Beton Numune ve Kırım Takibi · Kod Bloğu

> Şantiye › Saha Mühendisi · Workers / Workless

Beton dökümlerinin **7 ve 28 günlük numune kırım sonuçlarını** toplar; her 28 günlük sonucu ve ardışık sonuç gruplarını
**sınıf dayanımına göre** değerlendirir, sonucu gelmemiş dökümleri ve numuneler arası büyük farkları işaretler.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Karakteristik dayanım (fck) | Beton sınıfından: C30/37 → silindir numunede 30, küp numunede 37 MPa |
| Deney sonucu | Aynı döküm ve aynı yaştaki numunelerin ortalaması |
| Bireysel kriter | 28 günlük her sonuç ≥ fck − `--eksi` (varsayılan 4 MPa) |
| Grup kriteri | Aynı sınıf ve numune tipindeki ardışık `--grup` (varsayılan 3) sonucun ortalaması ≥ fck + `--ek` (varsayılan 4 MPa); gruplar örtüşmez, döküm tarihine göre |
| 7 gün oranı | 7 günlük sonuç ÷ fck × 100 (yalnız bilgi; uygunluk 28 günlük sonuçla değerlendirilir) |
| Numune farkı | (En büyük − en küçük) ÷ ortalama × 100; `--fark` (varsayılan %15) üstü bilgi notu |
| Eksik sonuç | Dökümün üzerinden 28 + `--tolerans` (varsayılan 2) gün geçmiş, 28 günlük sonuç yok |

Varsayılan değerler TS EN 206'daki **başlangıç üretimi** uygunluk kriteridir (3 sonuç: ortalama ≥ fck + 4, her sonuç
≥ fck − 4). Sürekli üretim ve farklı numune sayıları için kriter değişir; `--ek`, `--eksi` ve `--grup` ile projede
geçerli kriteri girin.

**Örnek:** C30/37 küp numunelerde fck = 37 MPa. BD-104, BD-105 ve BD-106'nın 28 günlük sonuçları 33,0, 39,9 ve
38,1 MPa → hepsi 33 MPa'nın (37 − 4) üstünde; ancak ortalama **37,0 MPa < 41 MPa** (37 + 4) → grup uygun değil.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                              # örnek veri (15.10.2026)
python main.py --numuneler beton_numuneleri.xlsx --tarih 15.10.2026
python main.py --numuneler beton_numuneleri.xlsx --ek 1,48 --grup 15 --fark 10
```

Her satır bir numunedir. Zorunlu sütunlar: Döküm No, Döküm Tarihi, Beton Sınıfı. Henüz kırılmamış dökümler için kırım
yaşı ve dayanımı boş bir satır yazılabilir. Numune tipi boşsa küp (15 cm) sayılır.

## Çıktı

`cikti/beton_kirim_takibi.xlsx`:

- `Kırım Sonuçları`: döküm × yaş; numune sayısı, ortalama, fck'ye oran, numune farkı, bireysel kriter.
- `Grup Değerlendirmesi`: sınıf bazında ardışık grupların ortalaması ve durumu.
- `Uyarılar`.

## Dikkat

- **Uygunluk kararı yetkili kişilerindir.** Sonuç uygun çıkmazsa yapı denetim kuruluşu, proje müellifi ve şantiye şefiyle
  birlikte değerlendirme yapılır (karot, tahribatsız deney, statik kontrol); kod yalnız işaretler.
- **Standardın güncel sürümünü kontrol edin.** Uygunluk kriterleri, numune alma sıklığı ve numune boyutu ilgili
  standart, yönetmelik ve sözleşme şartnamesine göre belirlenir; burada kod yalnız verdiğiniz sayıları uygular.
- **Laboratuvar raporu esastır.** Bu tablo laboratuvar deney raporlarının yerine geçmez; sonuçları rapordan aynen
  aktarın.
- **Numune farkı eşiği** laboratuvar ve şartnameye göre değişebilir; varsayılan %15 bilgi amaçlıdır.
- **Örnek veri:** Döküm numaraları, elemanlar ve dayanımlar kurgusaldır.

## Testler

- Sınıf okuma (küp / silindir fck), deney sonucu ortalaması, 7 gün oranı, numune farkı.
- Bireysel ve grup kriteri, eksik grup, eksik 28 günlük sonuç, seçenekler, okuma ve komut satırı.

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

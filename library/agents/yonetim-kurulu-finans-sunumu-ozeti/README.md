# Yönetim Kurulu Finans Sunumu Özeti · AI Agent

> Finans › Finans Müdürü · Gerisi Bende

Dönem sonuç tablosundan **yönetim kurulu için kısa, yorumlu finans özeti** hazırlar. Değişimleri, bütçe sapmalarını ve
oranları **kod hesaplar**; model yorumlar ve önceliklendirir. Metindeki her rakam kodun hesapladığı değerlerle
karşılaştırılır.

## Nasıl çalışır?

1. **Hesap (kod):**
   - Her kalemin önceki döneme göre değişimi (tutar, %) ve bütçe sapması (tutar, %). Eksi ya da parantezli yazılan gider
     kalemleri mutlak değerle değerlendirilir ("gider arttı").
   - Tanınan kalemlerden brüt kâr, FAVÖK ve net kâr marjı (puan farkıyla) ve cari oran.
   - Bütçeden en çok sapan 3 kalem.
   - Yönetim notları N1, N2 … diye numaralanır.
2. **Özet (model):** Öne çıkanlar, bölümler (gelir, kârlılık, bütçe, nakit ve borç, riskler) ve karar gerektiren konular
   yazılır. Her madde dayandığı kalem / oran / not numarasını gösterir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Maddedeki bir sayı, dayandığı kalem / oran / notun hazır rakamlarından biri değil (yuvarlama ya da yeni hesap)
   - "Arttı / azaldı" yönü değişimin işaretiyle çelişiyor
   - Neden-sonuç ifadesi ("nedeniyle", "kaynaklı") bir yönetim notuna dayanmıyor
   - Bütçeden en çok sapan kalemlerden biri ele alınmamış
   - "Kesinlikle", "garanti" gibi ifadeler var

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek çeyrek sonuçları
python agent.py --girdi sonuclar.xlsx --notlar notlar.txt --donem "2026 3. Çeyrek" --birim "milyon TL"
```

Tablo örneği (Bütçe ve Önceki Dönem isteğe bağlı):

| Kalem | Cari Dönem | Önceki Dönem | Bütçe |
|---|---|---|---|
| Hasılat | 1.250,0 | 1.080,0 | 1.300,0 |
| Satışların Maliyeti | (905,0) | (760,0) | (930,0) |
| Net Dönem Kârı | 71,0 | 86,0 | 105,0 |

Oranlar için kalem adları: Hasılat / Net Satışlar, Brüt Kâr, FAVÖK / EBITDA, Net Dönem Kârı, Dönen Varlıklar, Kısa Vadeli
Yükümlülükler.

## Çıktı

- `cikti/yk_finans_ozeti.xlsx`:
  - `Özet`: maddeler ve dayanakları.
  - `Rakamlar`: kodun hesapladığı değişimler, bütçe sapmaları ve oranlar.
  - `Yönetim Notları`.
- `cikti/yk_finans_ozeti.txt`: sunuma ya da e-postaya eklenecek metin.

## Dikkat

- **Taslaktır.** Finans müdürü onaylamadan yönetim kuruluna sunulmamalıdır. Rakamlar girdiğiniz tabloya dayanır;
  tablonun doğruluğu (konsolidasyon, TFRS / VUK farkları) sizin sorumluluğunuzdadır.
- **Neden-sonuç.** Model, yönetim notunda yazmayan bir nedeni yorum olarak eklemez; notlarınız ne kadar açıksa özet o
  kadar iyi olur.
- **Gizlilik.** Dönem rakamları ve notlar modele gönderilir. Kamuya açıklanmamış finansal bilgi için şirketinizin
  içeriden öğrenilen bilgi politikasına (halka açık şirketlerde SPK düzenlemeleri) uyun.
- **Örnek veri:** Rakamlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek:
- hasılat artışını "%16" diye yuvarlar,
- azalan net kâra "yükseldi" der,
- finansman giderindeki artışın nedenini nota dayandırmadan yazar,
- bütçeden çok sapan esas faaliyet kârını atlar,
- "kesinlikle" der.

Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 213 sayılı Vergi Usul Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.4.213.pdf) | Resmî metinle karşılaştırılan maddeler: 274, 280, 315, mükerrer 315, 320, 323 |

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

# Hasta Kabul Gün Sonu Kasa Kontrolü · Kod Bloğu

> Hasta Hizmetleri › Hasta Kayıt Kabul Görevlisi · Workers / Workless

HBYS'den alınan **kasa hareketlerini** (tahsilat, iade, iptal) sayılan nakit, POS gün sonu ve günün faturalarıyla
karşılaştırır; farkları, faturasız tahsilatları ve şüpheli iadeleri işaretler, imzaya hazır **kasa devir raporu**
hazırlar. İnternete ve HBYS'ye bağlanmaz; dışa aktarılmış dosyaları okur.

## Nasıl hesaplar?

| Kalem | Formül |
|---|---|
| Beklenen nakit | Önceki günden devir + nakit tahsilat − nakit iade − bankaya / merkez kasaya teslim |
| Nakit farkı | Sayılan nakit − beklenen nakit |
| HBYS kart neti | Kartla tahsilat − karta iade |
| Kart farkı | POS gün sonu toplamı − HBYS kart neti |
| Sonraki güne devir | Sayılan nakit (sayım girilmemişse beklenen nakit) |
| Fatura farkı | Faturadaki hasta payı − o faturaya yapılan net tahsilat (tahsilat − iade) |

İptal edilen makbuzlar toplamlara girmez, ayrıca listelenir. Ödeme türü "nakit", "kart / kredi / POS", "havale / EFT"
sözcüklerinden tanınır; diğerleri kendi adıyla ayrı satırda toplanır.

**Kontroller:** nakit ve kart farkı (`--tolerans`, varsayılan 1 TL üstü) · aynı makbuz numarası iki kez · fatura
numarası olmayan tahsilat · fatura listesinde olmayan fatura numarası · faturanın eksik ya da fazla tahsilatı · aynı
protokole o gün yapılan iadenin tahsilattan fazla olması · iptal edilen makbuzlar.

**Örnek:** Devir 2.500 + nakit tahsilat 14.320 − nakit iade 150 − bankaya teslim 10.000 = beklenen **6.670**; sayılan
6.665 → fark **−5** → "Kasa eksik".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                          # örnek veri
python main.py --hareketler kasa_hareketleri.xlsx --faturalar gunun_faturalari.xlsx --devir 2500 --bankaya 10000 --sayim 6665 --pos 19450
python main.py --hareketler kasa_hareketleri.xlsx --sayim 6665 --tolerans 0 --gun 30.09.2026
```

Zorunlu sütunlar: hareketlerde Makbuz No ve Tutar (İşlem boşsa tahsilat, Ödeme Türü boşsa nakit sayılır);
faturalarda Fatura No ve Tahsil Edilecek (hasta payı). İade tutarı eksi ya da artı yazılabilir.

## Çıktı

`cikti/hasta_kabul_kasa_devir.xlsx`:

- `Kasa Devir Raporu`: devir → tahsilat → iade → teslim → beklenen → sayım → fark; kart ve POS; devreden nakit;
  teslim eden / teslim alan imza satırları.
- `Ödeme Türleri`, `Kullanıcılar` (kullanıcı başına tahsilat, iade, iptal, nakit net), `Fatura Kontrolü`,
  `Hareketler`.
- `Uyarılar`.

## Dikkat

- **Kurum prosedürü esastır.** Kasa devir, bankaya yatırma, iade ve iptal onayı süreçleri hastaneden hastaneye
  değişir; rapor bu süreçlerin yerine geçmez, kontrolü kolaylaştırır.
- **Fark tespiti suistimal tespiti değildir.** Fark ve iade uyarıları inceleme önerisidir; kayıt hatası, geç girilen
  işlem ya da önceki güne ait iade olabilir.
- **Kişisel ve sağlık verisi içerir.** Hasta adı ve protokol numarası KVKK kapsamında özel nitelikli veriyle
  ilişkilidir; dosyayı yalnız yetkili kişilerle paylaşın, mümkünse hasta adı sütununu çıkarın.
- **HBYS dışa aktarım biçimi** kuruma göre farklıdır; sütun adları tanınmazsa başlıkları README'deki adlarla
  değiştirin.
- **Örnek veri:** Makbuz, protokol, hasta baş harfleri ve tutarlar kurgusaldır.

## Testler

- Ödeme türü toplamları, beklenen nakit, nakit ve POS farkı, devreden nakit, kullanıcı özeti.
- Fatura kontrolü, faturasız tahsilat, iadenin tahsilatı aşması, iptal, mükerrer makbuz, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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

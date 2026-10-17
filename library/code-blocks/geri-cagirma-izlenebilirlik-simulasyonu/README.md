# Geri Çağırma (İzlenebilirlik) Simülasyonu · Kod Bloğu

> Kalite Güvence ve Gıda Güvenliği › Kalite Güvence Uzmanı · Gerisi Bende

Şüpheli bir **hammadde lotundan** veya **üretim partisinden** yola çıkar; etkilenen üretim partilerini, bu partilerin
gittiği müşterileri ve miktarları çıkarır. Geri çağırma tatbikatında "bir adım geri, bir adım ileri" izlemeyi
dakikalar içinde yapmanızı sağlar. İnternete bağlanmaz.

## Nasıl çalışır?

| Adım | Kural |
|---|---|
| İleri izleme (`--lot`) | Lotu kullanan üretim partileri; o parti başka partide yarı mamul olarak kullanıldıysa o partiler de |
| Geri izleme (`--parti`) | Partinin kullandığı hammadde lotları (yarı mamul üzerinden gelenler dahil) ve partiyi kullanan partiler |
| Sevk edilen | Etkilenen partinin sevkiyat kayıtları toplamı |
| İç kullanım | Partinin başka üretim partilerinde yarı mamul olarak kullanılan miktarı |
| Depoda olması gereken | Üretilen − sevk edilen − iç kullanım |

Yarı mamul bağlantısı, kullanım kaydındaki "Hammadde Lot" değerinin bir üretim partisi numarasına eşit olmasıyla
kurulur. Seviye 1 lotu doğrudan kullanan, seviye 2 yarı mamul üzerinden etkilenen partidir.

**Örnek:** HM-2609-A salça lotu YM-0501 sos bazına (900 kg) ve UP-1003 partisine girmiş; YM-0501 de UP-1001 ve
UP-1004 partilerinde kullanılmış. Sonuç: 4 parti, 3 müşteri, 6.450 birim sevk edilmiş.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, HM-2609-A lotu
python main.py --uretim uretim.xlsx --kullanim hammadde_kullanim.xlsx --sevkiyat sevkiyat.xlsx --lot HM-2609-A
python main.py --uretim uretim.xlsx --kullanim hammadde_kullanim.xlsx --sevkiyat sevkiyat.xlsx --parti UP-1001
```

- Üretim: Üretim Partisi, Ürün, Üretim Tarihi, Üretilen Miktar, Birim.
- Hammadde kullanımı: Üretim Partisi, Hammadde, Hammadde Lot, Kullanılan Miktar.
- Sevkiyat: Üretim Partisi, Müşteri, Sevk Tarihi, Miktar, İrsaliye No.

Miktarlarda binlik ayırıcı kullanmayın (`4000`, `12,5`).

## Çıktı

`cikti/geri_cagirma_simulasyonu.xlsx`:

- `Özet`: etkilenen parti ve müşteri sayısı; üretilen, sevk edilen, depoda olması gereken miktar.
- `Etkilenen Partiler`: parti bazında miktarlar ve sayım sonucu için boş sütun.
- `Müşteriler`: müşteri bazında toplam miktar, partiler, ilk ve son sevk; bildirim durumu için boş sütun.
- `Sevkiyatlar`: etkilenen partilerin irsaliye bazında dökümü.
- `Lot Bağlantıları`: hangi lotun hangi partiye ne kadar girdiği.
- `Uyarılar`: sevk üretimden fazla, üretim kaydı yok, okunamayan satırlar.

## Dikkat

- **Sonuç, kayıtlarınız kadar doğrudur.** Kayda geçmemiş bir lot değişimi, ortak tank veya silo, yeniden işleme
  (rework) ya da elle düzeltilmiş parti numarası etkilenen ürünü gizler. Simülasyonun amacı bu boşlukları bulmaktır.
- **"Depoda olması gereken" hesaptır.** Fiili miktarı sayımla doğrulayın; fark varsa izlenebilirlik kaydınız eksiktir.
- **Aynı hammadde lotunu paylaşan hat ve ekipman kapsam dışıdır.** Çapraz bulaşma riski olan komşu partileri siz
  değerlendirirsiniz.
- **Müşteriden sonrası izlenmez.** Distribütörün ürünü kime sattığı bu kayıtlarda yoktur.
- **Karar aracı değildir.** Gerçek bir geri çağırma veya geri toplama kararı, yetkili makama bildirim ve kamuoyu
  duyurusu yürürlükteki gıda mevzuatına göre, işletmenizin yetkili kişilerince yapılır.
- **Örnek veri:** Ürünler, lotlar, partiler ve müşteriler kurgusaldır.

## Testler

- İleri izleme (yarı mamul zinciri dahil), müşteri toplamları, miktar dengesi, sevk üretimden fazla uyarısı.
- Geri izleme, etkilenmeyen partinin dışarıda kalması, döngülü kayıtta sonsuz döngüye girmeme.
- Bulunamayan lot, eksik üretim kaydı, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

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

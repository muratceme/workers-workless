# Satın Alma Tasarruf Raporu · Kod Bloğu

> Satın Alma › Satın Alma Müdürü · Gerisi Bende

Gerçekleşen alımları bütçe fiyatıyla karşılaştırır ve satın almanın sağladığı tasarrufu **kategori, tasarruf türü,
satın almacı ve ay bazında** raporlar. Aşımları tasarruftan ayrı gösterir; böylece iyi sonuç kötüyü gizlemez.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Bütçe tutarı, gerçekleşen tutar | Miktar × bütçe fiyatı; miktar × gerçekleşen fiyat |
| Tasarruf | Bütçe tutarı − gerçekleşen tutar. Eksi değer bütçe aşımıdır |
| Tasarruf oranı | Tasarruf ÷ bütçe tutarı |
| Brüt tasarruf, aşım, net | Artı tasarrufların toplamı; eksi olanların toplamı; ikisinin farkı |
| Pazarlık kazancı | (İlk teklif fiyatı − gerçekleşen fiyat) × miktar. Ayrı bir ölçüdür, tasarrufa eklenmez |

**Örnek:** 1.000 adet, bütçe fiyatı 75, ilk teklif 72, gerçekleşen 64 → tasarruf (75 − 64) × 1.000 = **11.000**
(%14,67); pazarlık kazancı (72 − 64) × 1.000 = **8.000**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                  # örnek veri
python main.py --alimlar alimlar.xlsx
python main.py --alimlar alimlar.xlsx --asim 3 --supheli 30
```

Zorunlu sütunlar: Malzeme, Miktar, Gerçekleşen Fiyat. Diğerleri: Tarih, Kategori, Tedarikçi, Bütçe Fiyatı, İlk
Teklif Fiyatı, Tasarruf Türü (Pazarlık, Alternatif Tedarikçi, Toplu Alım gibi; siz adlandırırsınız), Satın Almacı.

| Seçenek | Anlamı | Varsayılan |
|---|---|---|
| `--asim` | Bütçe aşımı bu yüzde ve üstündeyse uyarılır | 5 |
| `--supheli` | Tasarruf oranı bu yüzde ve üstündeyse bütçe fiyatı sorgulanır | 40 |

## Çıktı

`cikti/satin_alma_tasarruf.xlsx`:

- `Özet`: toplam harcama, bütçe tutarı, brüt tasarruf, aşım, net tasarruf ve oranı, pazarlık kazancı.
- `Kategori`, `Tasarruf Türü`, `Satın Almacı`, `Ay`: aynı göstergelerin kırılımları.
- `Alımlar`: satır bazında bütçe ve gerçekleşen tutar, tasarruf, oran, pazarlık kazancı.
- `Uyarılar`: bütçe aşımları, bütçe fiyatı olmayan alımlar, olağandışı yüksek tasarruf oranları, ilk tekliften
  pahalı alımlar.

## Dikkat

- **Tasarruf, bütçe fiyatı kadar gerçektir.** Bütçe fiyatı yüksek belirlendiyse tasarruf kâğıt üzerinde kalır.
  Bütçe fiyatının nasıl belirlendiğini (geçen yıl fiyatı, enflasyon varsayımı, piyasa araştırması) yazılı hale
  getirin ve yıl içinde değiştirmeyin.
- **Bu rapor maliyetten kaçınmayı ölçer.** Bütçeye göre tasarruf, geçen yıla göre fiyat düşüşü ile aynı şey
  değildir ve kâr-zarar tablosunda doğrudan görünmeyebilir. Finansla aynı tanımı kullandığınızdan emin olun.
- **Tek para birimi.** Fiyatları aynı para biriminde girin; dövizli alımlarda kur farkı tasarruf ya da aşım gibi
  görünür. Kur etkisini ayırmak isterseniz bütçe fiyatını da aynı kurla çevirin.
- **Toplam sahip olma maliyeti yoktur.** Vade, nakliye, kalite, stok ve garanti koşullarındaki fark hesaba
  girmez; ucuz fiyat her zaman ucuz alım değildir.
- **Kişi bazlı kırılım:** Satın almacı karşılaştırması, baktıkları kategoriler farklıysa adil değildir;
  performans değerlendirmesinde tek başına kullanmayın.
- **Örnek veri:** Malzemeler, tedarikçiler, kişiler ve fiyatlar kurgusaldır.

## Testler

- Tasarruf, oran, brüt, aşım, net ve pazarlık kazancı (elle hesaplanmış); bütçe fiyatı olmayan alım.
- Kategori, tür, satın almacı ve ay kırılımları; aşım ve yüksek oran eşik sınırları.
- Okuma hataları ve komut satırı.

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

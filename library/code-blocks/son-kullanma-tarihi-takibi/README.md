# Son Kullanma Tarihi Takibi · Kod Bloğu

> Eczane ve Tıbbi Malzeme › Hastane Eczacısı · Gerisi Bende

İlaç ve sarf stoklarını **lot bazında son kullanma tarihine göre** listeler; süresi geçmiş ve yaklaşan lotları,
**tüketim hızına göre süresi içinde tükenmeyecek miktarı** ve bunun tutarını hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Son kullanma tarihi | GG.AA.YYYY ya da yalnız ay-yıl (AA/YYYY, YYYY-AA); ay-yıl yazılırsa ayın son günü |
| Günlük tüketim | Günlük Tüketim sütunu · ya da Aylık Tüketim ÷ 30 · ya da `--tuketim` dosyasındaki dönem tüketimi ÷ `--gun` |
| Tüketim sırası | Aynı malzemenin lotları son kullanma tarihi en yakın olandan başlayarak tüketilir (FEFO) |
| Süresinde kullanılabilecek | Günlük tüketim × kalan gün − kendisinden önce tüketilecek lotların miktarı (0 ile lot miktarı arasında, aşağı yuvarlanır) |
| Risk miktarı | Lot miktarı − süresinde kullanılabilecek miktar |
| Risk tutarı | Risk miktarı × birim fiyat |

| Durum | Kural |
|---|---|
| Süresi geçmiş | Son kullanma tarihi rapor tarihinden önce (tüm miktar risk) |
| Tükenmeyecek | Risk miktarı sıfırdan büyük |
| Yaklaşıyor | Kalan gün `--yakin` (varsayılan 90) ve altı |
| Normal | Diğerleri |

**Örnek:** Seftriakson aylık 150 flakon → günde 5. 31.10.2026 SKT'li lotta 180 flakon var, 31 gün kalmış → 31 × 5 = 155
flakon kullanılabilir, **25 flakon** tükenmez → 25 × 38,50 = **962,50 TL** risk.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                     # örnek veri (30.09.2026)
python main.py --stok ilac_stok_lotlari.xlsx --tarih 30.09.2026 --yakin 120
python main.py --stok ilac_stok_lotlari.xlsx --tuketim son_3_ay_tuketim.xlsx --gun 90
```

Zorunlu sütunlar: Malzeme Kodu, Son Kullanma Tarihi, Miktar. Tüketim bilgisi olmayan malzemelerde yalnız tarih
kontrolü yapılır. Miktarı sıfır olan lotlar atlanır.

## Çıktı

`cikti/son_kullanma_takibi.xlsx`:

- `Lotlar`: son kullanma tarihine göre sıralı; kalan gün, durum (renkli), süresinde kullanılabilecek ve risk miktarı,
  risk tutarı.
- `Malzemeler`: risk tutarına göre sıralı; toplam miktar, günlük tüketim, stok günü (süresi geçmişler hariç), en yakın
  son kullanma tarihi.
- `Uyarılar`.

## Dikkat

- **Klinik ve mevzuat kararı eczacınındır.** Süresi geçmiş ilacın kullanımdan kaldırılması, imhası, iade ve transfer
  işlemleri kurum prosedürüne ve ilgili mevzuata göre yapılır; kod yalnız listeler.
- **Tüketim tahmini doğrusaldır.** Mevsimsel ilaçlar (ör. grip dönemi), salgın, yeni protokol ya da servis kapanışı
  tüketimi değiştirir. Tüketim hızını güncel dönemden alın.
- **FEFO varsayımı:** Hesap, lotların son kullanma sırasıyla tüketildiğini varsayar. Farklı depolardaki lotlar aynı
  havuzdan tüketiliyor sayılır; depolar arası transfer gerekmeden tüketilemeyecek stok daha erken risk oluşturabilir.
- **Açılmış çok dozluk flakonlar, soğuk zincir ve sulandırılmış ürünler** için açıldıktan sonraki kullanım süresi
  ayrıca takip edilmelidir; bu hesapta yoktur.
- **Örnek veri:** Malzeme kodları, lotlar, tarihler ve fiyatlar kurgusaldır.

## Testler

- FEFO sırası, süresinde kullanılabilecek ve risk miktarı, süresi geçmiş lot, yaklaşan, tüketim bilgisi olmayan malzeme.
- Ay-yıl SKT okuma, dış tüketim dosyası, malzeme özeti ve stok günü, okuma ve komut satırı.

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

# Araç Stok Yaşlandırma · Kod Bloğu

> Satış › Satış Müdürü · Workers / Workless

Bayi stoğundaki araçları **stoğa giriş tarihine göre yaşlandırır**; model, renk ve yaş bandı bazında adet, alış değeri,
ortalama stok günü ve **finansman maliyetini** hesaplar. Uzun bekleyen ve finansman maliyeti brüt marjı eriten araçları
işaretler. İnternete ve DMS'e bağlanmaz; dışa aktarılmış stok listesini okur.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Stok günü | Rapor tarihi − stoğa giriş tarihi (satılmış / teslim edilmiş araçlar alınmaz; rezerve araçlar alınır ve ayrıca sayılır) |
| Yaş bandı | 0–30 · 31–60 · 61–90 · 91–180 · 180+ gün |
| Faizli gün | max(stok günü − faizsiz gün, 0) |
| Finansman maliyeti | Alış maliyeti × yıllık oran ÷ 100 × faizli gün ÷ 365 (basit faiz) |
| Oran | Floor plan / kredi ile alınanlara `--faiz`; Finansman = Özkaynak / Nakit / Peşin olanlara `--firsat-faiz` (fırsat maliyeti) |
| Marj tüketimi % | Finansman maliyeti ÷ (liste fiyatı − alış maliyeti) × 100 |

**Uyarılar:** 180 günü geçen araç (Yüksek) · 90 günü geçen araç · finansman maliyeti brüt marjın `--marj-esik` (varsayılan %50)
kadarını aştı (Orta) · faizsiz süre 7 gün içinde bitiyor · liste fiyatı maliyetin altında · önceki model yılı · şasi numarası
17 karakter değil ya da I / O / Q içeriyor · faiz oranı verilmedi (Bilgi) · aynı şasi iki kez (Yüksek).

**Örnek:** 1.000.000 TL'ye alınan, 60 gün faizsiz floor plan ile 150 gündür stokta bekleyen araç: 1.000.000 × %36 × 90 ÷ 365
= **88.767,12 TL** finansman maliyeti; 80.000 TL brüt marjın **%110,96**'sı.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                              # örnek veri (14.10.2026, faiz %36)
python main.py --stok stok.xlsx --faiz 36
python main.py --stok stok.xlsx --tarih 31.10.2026 --faiz 36 --firsat-faiz 25 --marj-esik 40
```

Zorunlu sütunlar: Şasi No, Stoğa Giriş Tarihi ve Alış Maliyeti. `--faiz` verilmezse maliyet hesaplanmaz; yalnız yaşlandırma
yapılır.

## Çıktı

`cikti/arac_stok_yaslandirma.xlsx`: `Yaş Bandı` (toplam satırıyla), `Model`, `Renk`, `Araç Listesi` (en eski üstte; 90+ gün sarı,
180+ gün kırmızı), `Uyarılar`.

## Dikkat

- **Faiz oranı sizin girdinizdir.** Örnekteki %36 yalnız örnektir; floor plan / kredi sözleşmenizdeki oranı, faizsiz süreyi ve
  masrafları kullanın. Gerçek faiz bileşik, değişken ya da komisyonlu olabilir; ekstreyle karşılaştırın.
- **Brüt marj liste fiyatına göredir.** Kampanya indirimleri, marka primleri ve hedef bonusları hesapta yoktur.
- **Fırsat maliyeti** gerçek bir gider değildir; özkaynakla alınan aracın bağladığı paranın alternatif getirisini gösterir.
- **Örnek veri:** Marka, modeller, şasi numaraları ve fiyatlar kurgusaldır.

## Testler

- Stok günü, yaş bandı, faizsiz gün, floor plan ve özkaynak oranıyla finansman maliyeti, marj tüketimi, sıralama ve özetler.
- Yaş, marj, faizsiz süre, model yılı ve şasi biçimi uyarıları; durum tanıma, okuma ve komut satırı.

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

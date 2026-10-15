# Yönetim Raporu (KPI) · Kod Bloğu

> Finans › Bütçe ve Raporlama Uzmanı · Workers / Workless

Aylık yönetim raporunun **tablo ve grafiklerini** satış verisinden ve isteğe bağlı mizandan otomatik hazırlar.
Her ay elle güncellenen Excel sunumunun yerine geçer. İnternete bağlanmaz.

## Göstergeler

Her gösterge 5 sütunda karşılaştırılır: **rapor ayı**, önceki ay, **geçen yıl aynı ay**, yılbaşından bugüne
(**YTD**) ve geçen yıl YTD.

| Satış | Finans (mizan verilirse) |
|---|---|
| Net satış ve büyüme | Net satış, brüt kâr, esas faaliyet kârı, dönem net kârı |
| Satış adedi, fatura sayısı, ortalama fatura tutarı | Hazır değerler, ticari alacak / borç, stok |
| Aktif ve yeni müşteri | Cari oran, kaldıraç, brüt ve net kâr marjı |
| Brüt kâr ve marjı (maliyet verilmişse) | |
| İlk 5 müşterinin payı (yoğunlaşma riski) | |

- **Kırılımlar ve grafikler:** Kanal, kategori, müşteri ve ürün kırılımları (GY ve YTD karşılaştırmalı), 13 aylık
  satış ve marj trendi.
- **İadeler:** `Tür` sütunu "İade" olan satırlar pozitif yazılmışsa ters çevrilip satıştan düşülür.
- **Oranlar:** Oran değişimleri **yüzde puan** olarak verilir.
- **Mizan tarafı:** **Mali Tablo Rasyo Analizi** kod bloğunun çekirdeği kullanılır (`rasyo_cekirdek.py`).

> "Yeni müşteri", veri setinde ilk alımı o dönemde olan müşteridir. Verinin başladığı ilk aylarda bu sayı yüksek
> görünür; en az 12 aylık geçmiş verin.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                          # örnek 21 aylık satış verisi (Eylül 2026)
python main.py --satis satislar.xlsx --ay 2026-09
python main.py --satis satislar.xlsx --ay 2026-09 --mizan mizan_eylul.xlsx
```

## Çıktı

`Yönetim Özeti` (KPI tablosu + grafikler) · `Trend` · `Kanal` · `Kategori` · `Müşteri` · `Ürün`

## Testler

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

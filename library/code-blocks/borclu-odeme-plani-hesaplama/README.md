# Borçlu Ödeme Planı Hesaplama · Kod Bloğu

> İcra Takip › İcra Takip Elemanı (Avukat Katibi) · Workers / Workless

İcra dosyasındaki toplam borcu, peşinatı, taksit sayısını ve faiz oranını alıp **faiz dahil ödeme planı** ve ödeme protokolüne
konacak **özet metni** çıkarır. Tek dosya için komut satırından, çok dosya için tablodan çalışır. İnternete ve UYAP'a bağlanmaz.

## Nasıl hesaplar?

| Yöntem | Formül |
|---|---|
| Taksitlenen anapara | Toplam borç − peşinat |
| Azalan bakiye (varsayılan) | Her taksitte eşit anapara + kalan anapara × yıllık faiz × gün ÷ 365. Gün: protokol tarihinden ilk taksite, sonra taksitten taksite |
| Eşit taksit (`--yontem esit`) | Aylık oran r = yıllık ÷ 12; taksit = A × r ÷ (1 − (1 + r)^−n); faiz = kalan × r |
| Vadeler | İlk taksit tarihinden itibaren aylık; ay sonuna denk gelen günler kısa aylarda ayın son gününe çekilir |

Son taksit, kuruş farkları kapanacak şekilde kalan anaparanın tamamını içerir.

**Örnek:** 120.000 TL borç, 20.000 TL peşinat, 10 taksit, yıllık %24, protokol 14.10.2026, ilk taksit 15.11.2026 → ilk taksit
faizi 100.000 × %24 × 32 ÷ 365 = **2.104,11 TL**, ilk taksit **12.104,11 TL**, toplam faiz **11.059,73 TL**.

**Uyarılar:** plan çıkarılamayan dosya (peşinat ≥ borç, ilk taksit protokolden önce vb.) · okunamayan satır (Yüksek) · faizsiz plan ·
protokol tarihi yazılmamış (ilk taksitten 30 gün önce kabul edilir) · 36 taksitten uzun vade (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                            # örnek dosyalar
python main.py --borc 120.000 --pesinat 20.000 --taksit 10 --faiz 24 --ilk 15.11.2026 --protokol 14.10.2026 --borclu "Ad Soyad" --dosya-no "2026/1452 E."
python main.py --borc 48.500 --taksit 6 --faiz 24 --ilk 01.11.2026 --yontem esit
python main.py --dosyalar dosyalar.xlsx
```

## Çıktı

`cikti/odeme_plani.xlsx`: `Özet` (protokol özeti metniyle), `Ödeme Planı`, `Uyarılar`.

## Dikkat

- **Faiz oranı sizin girdinizdir.** Takip talebi, ilam ya da sözleşmedeki faiz türünü (yasal, ticari temerrüt, akdî) ve güncel
  oranı kullanın. Örnek dosyadaki %24 ve %30 yalnız örnektir. Değişken faizli alacaklarda plan, oran değiştiğinde yeniden
  hesaplanmalıdır.
- **Plan hukuki metin değildir.** İcra ve İflas Kanunu'ndaki taksit / ödeme sözü hükümleri, takibin durdurulup durdurulmayacağı,
  muacceliyet şartı, masraf ve vekâlet ücretinin plana dahil olup olmadığı protokolde avukat tarafından yazılmalıdır.
- **Toplam borç** dosya hesabındaki (UYAP) anapara, işlemiş faiz, masraf ve vekâlet ücreti toplamıyla aynı olmalıdır.
- **Örnek veri:** Dosya numaraları, borçlular ve tutarlar kurgusaldır.

## Testler

- Azalan bakiye (gün bazlı faiz), eşit taksit (annüite), son taksit kapanışı, ay sonu vadeleri, peşinat, toplamlar.
- Hatalı girişler, yöntem tanıma, uyarılar, protokol metni, okuma, tek ve toplu komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

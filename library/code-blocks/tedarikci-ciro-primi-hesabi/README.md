# Tedarikçi Ciro Primi Hesabı · Kod Bloğu

> Satın Alma ve Kategori Yönetimi › Kategori Yöneticisi · Workers / Workless

Tedarikçi anlaşmalarındaki **ciro primi kademelerini** dönem alımlarına uygular; hak edilen primi, sonraki kademeye
kalan ciroyu ve bu hızla **dönem sonunda** ulaşılacak primi hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Net ciro | Dönemdeki alımlar − iadeler (KDV hariç) |
| Tüm ciro yöntemi | Ulaşılan en yüksek kademenin oranı **tüm** net ciroya uygulanır |
| Dilimli yöntem | Her dilimdeki ciro kendi kademesinin oranıyla primlendirilir |
| Ulaşılırsa ek prim | Sonraki kademenin alt sınırındaki prim − bugünkü prim |
| Dönem sonu tahmini | Net ciro ÷ geçen gün × dönemin gün sayısı |

**Örnek (Tüm ciro):** Kademeler 0 → %1, 2.000.000 → %2, 3.500.000 → %3. Net ciro 2.577.700 → %2 kademesi →
prim 2.577.700 × %2 = **51.554**.

**Örnek (Dilimli):** Kademeler 0 → %0,5, 1.000.000 → %1,5, 2.000.000 → %2,5. Net ciro 1.568.100 → 1.000.000 × %0,5
+ 568.100 × %1,5 = 5.000 + 8.521,50 = **13.521,50**.

Dönem sonu tahmini yalnız `--baslangic`, `--bitis` ve `--tarih` verildiğinde ve tarih dönemin içindeyse hesaplanır.
Tarihi olan alımlardan dönem dışında kalanlar hesaba girmez.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                       # örnek veri (01.01–31.12.2026 dönemi, 30.09.2026 itibarıyla)
python main.py --anlasmalar ciro_primi_anlasmalari.xlsx --alimlar tedarikci_alimlari.xlsx
python main.py --anlasmalar ciro_primi_anlasmalari.xlsx --alimlar tedarikci_alimlari.xlsx --baslangic 01.01.2026 --bitis 31.12.2026 --tarih 30.09.2026
```

Anlaşma dosyasında her satır bir kademedir: Tedarikçi, Alt Sınır, Prim %, Yöntem ("Tüm ciro" ya da "Dilimli";
boşsa Tüm ciro). Koşulsuz prim için alt sınırı 0 yazın. Alım dosyasında iadeler eksi tutarla ya da Tür = "İade"
olarak yazılabilir.

## Çıktı

`cikti/tedarikci_ciro_primi.xlsx`:

- `Prim Hesabı`: alımlar, iadeler, net ciro, ulaşılan kademe, hak edilen prim, sonraki kademe, kalan ciro, ek prim,
  dönem sonu tahmini.
- `Kademeler`: okunan anlaşma kademeleri.
- `Aylık Ciro`: tedarikçi × ay net alım.
- `Uyarılar`: anlaşmasız tedarikçi, alım yok, sonraki kademeye yakın (%10 içinde) ya da ulaşılabilir, okunamayan
  satırlar.

## Dikkat

- **Ciro tanımı anlaşmaya bağlıdır.** Ciroya iadelerin, fatura altı iskontoların, fiyat farkı faturalarının ve
  promosyon ürünlerinin dahil olup olmadığını sözleşmeden kontrol edip alım dosyasını ona göre hazırlayın.
- **Muhasebe ve vergi işlemi ayrıdır.** Hesaplanan tutar KDV hariçtir. Prim faturasının kimin tarafından, hangi
  belgeyle ve hangi KDV uygulamasıyla düzenleneceğini mali müşavirinizle doğrulayın; kod fatura ya da kayıt üretmez.
- **Tahmin doğrusal bir uzatmadır.** Mevsimsellik, kampanya ve fiyat değişimlerini gözetmez; yalnız yön gösterir.
- **Hesap kontrol amaçlıdır.** Tedarikçinin mutabakatı ve sözleşmedeki ödeme koşulları esastır.
- **Örnek veri:** Tedarikçi adları, kademeler ve tutarlar kurgusaldır.

## Testler

- Tüm ciro ve dilimli yöntem, kademe sınırları, eşik altında prim olmaması, iade ve dönem dışı alım.
- Sonraki kademe ve ek prim, dönem sonu tahmini, uyarılar, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

# Ürün Başlığı SEO Kontrolü · Kod Bloğu

> Kategori ve İçerik Yönetimi › E-ticaret İçerik Uzmanı · Workers / Workless

Ürün başlıklarını **uzunluk, marka → ürün tipi → özellik sırası, eksik özellik, yasaklı ifade ve biçim** açısından
kontrol eder; her başlık için sorun listesi ve **önerilen başlık** üretir. İnternete bağlanmaz.

## Nasıl kontrol eder?

| Kontrol | Kural |
|---|---|
| Uzunluk | `--min` (varsayılan 30) ile `--max` (varsayılan 100) karakter arası |
| Marka | Başlıkta geçer ve başlık markayla başlar |
| Ürün tipi | Başlıkta geçer ve markadan sonra gelir |
| Özellik | Renk, Materyal, Beden / Ebat, Model sütunlarına yazılan değer başlıkta geçer (ek almış hâli de sayılır: Pamuk → Pamuklu) |
| Yasaklı ifade | Varsayılan liste ya da `--yasakli` dosyası; `--ek-yasakli` ile eklenir |
| Büyük harf | Harflerin yarısından fazlası büyük harf |
| Özel karakter | `! ? * # $ @ ~ ^ \| < > { } [ ] _ = +` ya da emoji |
| Tekrar | Üç harften uzun aynı kelime iki kez |
| Fazla boşluk | Çift boşluk, başta ya da sonda boşluk |
| Aynı başlık | Aynı başlık başka bir ürün kodunda da var |

Karşılaştırmalar Türkçe karakter ve büyük-küçük harf farkı gözetmeden yapılır.

**Önerilen başlık:** Sıra doğruysa başlığın kendi kelime sırası korunur; değilse **Marka + Ürün tipi + diğer
kelimeler** dizilir. Yasaklı ifadeler, özel karakterler ve tekrar eden kelimeler çıkarılır, tamamı büyük harf
yazılmış kelimeler düzeltilir, başlıkta olmayan özellikler sona eklenir.

**Varsayılan yasaklı ifadeler:** en ucuz, en iyi, en kaliteli, 1 numara, bir numara, indirim, indirimli, kampanya,
kampanyalı, ücretsiz kargo, bedava, fırsat, son fiyat, stokta, hemen al, acele, sınırlı stok, şok fiyat. Bu liste
bir başlangıçtır; pazaryerinizin içerik kurallarına göre `--yasakli` ile kendi listenizi verin.

**Örnek:** "Su Isıtıcı Lumora EN UCUZ!!! Paslanmaz Çelik" → marka başta değil, ürün tipi markadan önce, renk yok,
yasaklı ifade, özel karakter → öneri **"Lumora Su Isıtıcı Paslanmaz Çelik Gri"**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                  # örnek veri
python main.py --urunler urun_basliklari.xlsx --min 40 --max 120
python main.py --urunler urun_basliklari.xlsx --yasakli yasakli_ifadeler.xlsx --ek-yasakli "orijinal;garantili"
```

Zorunlu sütunlar: Ürün Kodu ve Başlık. Marka, Ürün Tipi ve özellik sütunları yazılmazsa ilgili kontrol yapılmaz.

## Çıktı

`cikti/urun_basligi_kontrolu.xlsx`:

- `Başlık Kontrolü`: mevcut başlık, uzunluk, sorun sayısı (renkli), sorunlar, önerilen başlık ve uzunluğu, aynı
  başlıklı ürünler. Öneri boşsa başlık değiştirilmeden kalabilir.
- `Sorun Özeti`: sorun türüne göre ürün sayısı.
- `Uyarılar`.

## Dikkat

- **Öneri otomatik üretilir.** Kelime sırası, anlam ve dil bilgisi kontrolü sizindir; özellikle "Ürün tipi markadan
  önce" düzeltmesinde niteleyici kelimeler ("Çelik Termos" → "Termos Çelik") doğal olmayan bir sıraya geçebilir.
- **Pazaryeri kuralları farklıdır.** Başlık uzunluğu sınırları, yasaklı kelimeler ve başlık yapısı pazaryerine ve
  kategoriye göre değişir; varsayılan değerler genel kabul görmüş sınırlar değildir.
- **Yanıltıcı ifadeler hukuki risk taşır.** "En ucuz", "en iyi" gibi üstünlük iddiaları ve gerçek olmayan indirim
  ifadeleri ticari reklam mevzuatı açısından sorun yaratabilir; kod yalnız listedeki ifadeleri arar, iddianın
  doğruluğunu değerlendirmez.
- **Başka markaların adı:** Başlıkta rakip ya da uyumlu ürün markası kullanmak marka hakkı sorununa yol açabilir; bu
  kontrol yapılmaz.
- **Örnek veri:** Markalar ve ürünler kurgusaldır.

## Testler

- Uzunluk, marka ve ürün tipi sırası, eksik özellik (ek almış hâli dahil), yasaklı ifade, büyük harf, özel
  karakter, tekrar, fazla boşluk, aynı başlık.
- Önerilen başlık (sıra korunarak ya da yeniden dizilerek), Türkçe büyük-küçük harf düzeltmesi, özel liste, okuma ve
  komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

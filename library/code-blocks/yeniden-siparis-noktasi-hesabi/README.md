# Yeniden Sipariş Noktası Hesabı · Kod Bloğu

> Üretim › Üretim Planlama · Workers / Workless

Her stok kalemi için **emniyet stoğunu**, **yeniden sipariş noktasını (YSN)** ve **ekonomik sipariş miktarını
(ESM)** hesaplar. Stok pozisyonu YSN'ye inmiş kalemler için sipariş önerir. İnternete bağlanmaz.

```
Emniyet stoğu = z × √( L × σd² + d² × σL² )
YSN           = d × L + Emniyet stoğu
ESM           = √( 2 × D × S / (h × c) )
Stok pozisyonu = eldeki + yoldaki (açık sipariş) − ayrılmış
```

**Semboller:**
- `d`, `σd`: günlük talep ve standart sapması
- `L`, `σL`: tedarik süresi (gün) ve standart sapması
- `z`: hizmet düzeyinin standart normal karşılığı (%95 → 1,645)
- `D`: yıllık talep
- `S`: sipariş başına maliyet
- `h`: yıllık stok tutma oranı
- `c`: birim maliyet

Formül iki tür belirsizliği birlikte hesaba katar: **talep** belirsizliği ve **tedarik süresi** belirsizliği.
Tedarik süresi sabitse (σL = 0) klasik `z × σd × √L` formülüne indirgenir.

## Neler yapar?

- **Talep kaynağı:** Talep kalem listesinde verilmemişse tüketim geçmişinden hesaplanır. Aylık toplamlar
  kullanılır: `d = ortalama ÷ 30`, `σd = σ ÷ √30` (`--donem-gun` ile değiştirilir).
- **Sipariş önerisi:** Pozisyon ≤ YSN ise ESM önerilir. ESM yoksa öneri, pozisyonu YSN + tedarik süresi talebine
  tamamlayacak miktardır. Öneri **en az sipariş miktarına** ve **ambalaj katına** yukarı yuvarlanır.
- **ACİL:** Eldeki stok, tedarik süresi dolmadan tükenecekse kalem acil olarak işaretlenir.
- **Emniyet stoğunun TL değeri:** Stokta bağlı kalan sermayeyi görmek için hesaplanır.

Formüller ders kitabı örnekleriyle test edilmiştir (`d = 20, σd = 4, L = 9, %95 → SS = 19,74`).

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                       # örnek kalemler + 12 aylık tüketim
python main.py --girdi kalemler.xlsx --gecmis tuketim.xlsx --donem-gun 30
python main.py --girdi kalemler.xlsx --hizmet 97,5 --siparis-maliyeti 750 --tutma-orani 25
```

> Talebi çok düzensiz kalemlerde normal dağılım varsayımı zayıflar ve emniyet stoğu hatalı çıkabilir. Bu
> kalemleri önce **Stok ABC-XYZ Analizi** ile ayırın (Z sınıfı).

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

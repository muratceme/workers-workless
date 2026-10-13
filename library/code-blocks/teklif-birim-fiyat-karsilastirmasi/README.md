# Teklif Birim Fiyat Karşılaştırması · Kod Bloğu

> İhale ve Teklif › İhale Mühendisi · Workers / Workless

Taşeron ve tedarikçi tekliflerini **poz bazında yan yana** getirir; firma toplamlarını ve her pozda en düşük fiyattan
oluşan **karma teklif cetvelini** çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Poz tutarı | Miktar × birim fiyat |
| En düşük | Pozdaki en düşük birim fiyat (eşitlikte adı alfabetik önce gelen firma) |
| Ortalamadan fark % | (Ortalama − en düşük) ÷ ortalama × 100 |
| Aşırı düşük / yüksek | En az 3 teklif varsa, medyanın `--dusuk` (%30) altı veya `--yuksek` (%30) üstü |
| Firma toplamı | Σ poz tutarı; yalnız bütün pozlara fiyat veren firmalar sıralanır |
| Karma cetvel toplamı | Σ miktar × pozun en düşük fiyatı |
| Karma tasarruf | En düşük tam teklif − karma cetvel toplamı |

**Örnek:** Saten boya, 5.200 m²; teklifler 95, 58 ve 92. Medyan 92; 58 medyanın %36,96 altında → "aşırı düşük".
En düşük tutar 5.200 × 58 = 301.600.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                   # örnek veri
python main.py --teklifler teklifler.xlsx
python main.py --teklifler teklifler.xlsx --dusuk 25 --yuksek 40
```

Dosyada her firma-poz bir satırdır: Poz No, İş Kalemi, Birim, Miktar, Firma, Birim Fiyat. Fiyat hücresi boş bırakılan
poz "eksik teklif" sayılır.

## Çıktı

`cikti/teklif_karsilastirma.xlsx`:

- `Karşılaştırma`: poz × firma birim fiyatları; en düşük yeşil, aykırı fiyatlar renkli.
- `Firma Toplamları`: teklif toplamı, eksik poz sayısı, sıra; karma cetveldeki payı.
- `Karma Cetvel`: her pozda en düşük fiyatı veren firma ve tutar.
- `Uyarılar`: aşırı düşük ve yüksek fiyat, eksik teklif, tek teklif, miktar tutarsızlığı, mükerrer teklif.

## Dikkat

- **Kapsam eşitliğini siz sağlarsınız.** Fiyatlar aynı para biriminde ve aynı kapsamda (KDV, nakliye, malzeme, iskele,
  montaj dahil veya hariç) olmalıdır. Kod kapsam farkını bilemez; elma ile armudu karşılaştırabilir.
- **En düşük fiyat en uygun teklif değildir.** Ödeme vadesi, iş bitirme, kalite, kapasite ve teminat koşulları ayrıca
  değerlendirilir. Aşırı düşük fiyat çoğu zaman eksik kapsam veya yazım hatasıdır.
- **Karma cetvel bir pazarlık aracıdır.** İşi pozlara bölüp farklı firmalara vermek her zaman mümkün değildir;
  firmalar tek poz için aynı fiyatı vermeyebilir.
- **Eksik teklif veren firmanın toplamı düşük görünür.** Bu yüzden yalnız tam teklifler sıralanır.
- **Kamu ihalesi değerlendirmesi değildir.** Aşırı düşük teklif sorgulaması ve sınır değer hesabı kamu ihale
  mevzuatındaki usule göre ayrıca yapılır.
- **Örnek veri:** Firmalar, pozlar ve fiyatlar kurgusaldır.

## Testler

- Poz bazında en düşük, ortalama, medyan; firma toplamları, tam teklif sırası, karma cetvel ve tasarruf.
- Aşırı düşük ve yüksek eşikleri, eşit fiyat, tek ve iki teklif, miktar tutarsızlığı, mükerrer teklif.
- Okunamayan satır, boş fiyat, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

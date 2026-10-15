# İade Nedeni Analizi · AI Agent

> E-ticaret Operasyon ve Kargo › E-ticaret Operasyon Uzmanı · Workers / Workless

İade açıklamalarını **neden kategorilerine** ayırır: beden küçük / büyük, kalıp, kusurlu ürün, yanlış gönderim, renk ya da
kalite beklenti farkı, geç teslimat, vazgeçme. Ardından ürün bazında iade oranını çıkarır ve her ürün için ilgili ekibe
aksiyon önerir.

## Nasıl çalışır?

1. **Hazırlık (kod):**
   - Açıklaması boş iadeler "Belirtilmemiş" sayılır ve modele gönderilmez.
   - Açık ifadeler kural işareti alır: "küçük / dar geldi", "büyük / bol geldi", "yanlış ürün / farklı beden geldi",
     "bozuk / sökük / yırtık / lekeli", "geç geldi".
2. **Sınıflandırma (model):** Her iade bir kategoriye ayrılır, dayanılan ifade alıntılanır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa iade modele geri gönderilir (varsayılan en fazla 2 tur):
   - Sınıflandırılmamış iade
   - Açıklamada olmayan alıntı
   - Tek bir kural işareti varken farklı kategori
4. **Analiz (kod):**
   - Ürün bazında iade oranı = iade / satış (satış dosyası verilirse). `--esik` %'yi aşanlar işaretlenir (varsayılan %15).
   - Beden yönü: küçük geldi iadeleri en az `--beden-min` ve büyük geldinin en az 2 katıysa "kalıbı dar" ürün sayfası notu
     önerilir (tersi için "kalıbı geniş").
   - Kusur ve yanlış gönderim → depo / tedarikçi aksiyonu.
   - Renk / görsel farkı → ürün fotoğrafı kontrolü.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                  # örnek iadeler + satış adetleri
python agent.py --girdi iadeler.xlsx --satislar satislar.xlsx
python agent.py --girdi iadeler.xlsx --satislar satislar.xlsx --esik 25 --beden-min 3
```

İade ve satış dosyaları aynı dönemi kapsamalıdır; aksi hâlde iade oranı yanıltıcı olur.

## Çıktı

- `cikti/iade_analizi.xlsx`:
  - `Ürün Analizi`: satış, iade, oran, kategori dağılımı ve önerilen aksiyonlar.
  - `İadeler`: kategori, alıntı ve kural işareti.
  - `Neden Dağılımı`: tabloyla birlikte grafik.

## Dikkat

- **Cayma hakkı.** Mesafeli satışlarda tüketici 14 gün içinde gerekçe göstermeden cayma hakkını kullanabilir (6502 sayılı
  Kanun ve Mesafeli Sözleşmeler Yönetmeliği). Bu yüzden açıklamasız ya da belirsiz iadeler normaldir. Bu analiz iade
  kabul kararı vermez; yalnız neden istatistiği üretir.
- **Eşik.** Kabul edilebilir iade oranı kategoriye göre çok değişir (giyim ve ayakkabıda yüksektir). `--esik` değerini kendi
  kategori ortalamanıza göre ayarlayın.
- **Kategoriler öneridir.** Kural işaretleri yalnız açık ifadeleri yakalar; diğerleri modelin yorumudur.
- **Veri gönderimi.** Yalnız iade açıklaması, ürün adı ve beden seçtiğiniz yapay zekâ sağlayıcısına gönderilir; sipariş no ve
  müşteri bilgisi gönderilmez. Açıklamadaki telefon ve e-posta maskelenir.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek "uzun geldi" iadesine "kalıp beğenilmedi", "geç geldi" iadesine
"vazgeçti" der ve açıklamada olmayan bir alıntı yazar. Kodun bunları yakalayıp yalnız bu iadeleri yeniden sorduğu, iade
oranlarını ve beden yönünü doğru hesapladığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve analizin doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.

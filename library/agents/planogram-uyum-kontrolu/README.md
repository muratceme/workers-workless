# Planogram Uyum Kontrolü · AI Agent

> Görsel Düzenleme › Görsel Düzenleme Uzmanı · Workers / Workless

Raf fotoğrafını planogram listesiyle karşılaştırıp **eksik, yanlış yerleştirilmiş, yüz sayısı eksik, fiyat etiketi
olmayan ya da etiket fiyatı farklı ürünleri** listeler. Görsel destekli model yalnız **fotoğrafta gördüğünü** kayda
geçirir. Uyum kararını model değil, kod planogramla karşılaştırarak verir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Planogram okunur; fotoğraflar uzun kenarı 1568 piksel olacak şekilde küçültülüp JPEG yapılır.
2. **Gözlem (model, görsel):** Her raf (1 = en üst) ve sıra (soldan) için ürün (planogramdaki barkod ya da
   `BILINMIYOR`), görünen yüz sayısı, fiyat etiketi (Var / Yok / Okunamıyor) ve etiketteki fiyat; boş alanlar; güven.
3. **Denetim (kod):** Aşağıdakilerden biri varsa gözlem modele geri gönderilir (varsayılan en fazla 2 tur):
   - Planogramda olmayan barkod (tahminle barkod yazılamaz; emin değilse `BILINMIYOR`)
   - Planogram aralığı dışında raf; aynı konuma iki ürün; hem ürün hem boş alan yazılmış konum
   - Etiket "Var" ama fiyat yok; gözlemi olmayan raf
4. **Karşılaştırma (kod):** Her planogram ürünü için Yerinde / Yanlış yer / Eksik; toplam yüz sayısı < planogram;
   etiket yok / okunamadı; etiket fiyatı ≠ planogram fiyatı. Düşük güvenli gözlemler "Elle Kontrol" diye işaretlenir.
   Planogramda olmayan ürünler ve boş alanlar ayrı sayfada listelenir; tam uyum oranı hesaplanır.

## Kurulum

```bash
pip install -r requirements.txt        # Pillow dahil
cp .env.example .env                   # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

Görsel destekleyen bir model gerekir (varsayılan Anthropic modeli görselleri okur). Ollama kullanıyorsanız görsel
destekli bir model seçin.

## Kullanım

```bash
python agent.py                                                    # örnek (kurgusal) sıcak içecek reyonu
python agent.py --planogram reyon12.xlsx --foto raf_ust.jpg raf_alt.jpg
```

**Fotoğraf ipuçları:** Rafı karşıdan, tamamı kadrajda ve fiyat etiketleri okunur şekilde çekin; yansıma ve bulanıklık
gözlem güvenini düşürür. Uzun reyonları birkaç fotoğrafa bölün.

## Çıktı

- `cikti/planogram_uyum.xlsx`: Uyum (yeşil: tam uyumlu, sarı: yerinde ama sorunlu, kırmızı: eksik / yanlış yer),
  Planogram Dışı.

## Dikkat

- **Fotoğraf yorumu hatalı olabilir.** Benzer ambalajlar, kısmen kapalı ürünler ve küçük etiket yazıları yanlış
  okunabilir. Kırmızı ve "Elle Kontrol" işaretli satırları rafta doğrulayın; fiyat düzeltmesini kasa / fiyat sisteminden
  kontrol ederek yapın.
- **Kişisel veri.** Fotoğraflar yapay zekâ sağlayıcısına gönderilir. Müşteri ya da çalışan yüzü, kartlı ödeme ekranı
  gibi bilgiler kadrajda olmamalıdır.
- **Örnek veri:** Ürünler, barkodlar ve fotoğraf (kodla çizilmiş) kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk gözlemde planogramda olmayan bir barkod uydurur, aynı konuma iki ürün yazar,
etiket "Var" deyip fiyatı boş bırakır, olmayan 4. rafa gözlem ekler ve 3. rafı atlar. Kodun beş sorunu da yakaladığı;
düzeltilmiş gözlemde yanlış yerdeki iki çayı, eksik hazır kahveyi, filtre kahvede yüz eksiğini ve 199,00 / 214,00 fiyat
farkını bulduğu doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve raf düzeltmeleri kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.

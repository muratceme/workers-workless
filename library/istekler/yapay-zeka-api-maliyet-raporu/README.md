# Yapay Zekâ API Maliyet Raporu · Kod Bloğu

> Sizden Gelenler › Abonelikli mobil uygulama · Gerisi Bende

Uygulamanızın yapay zekâ (LLM) API çağrılarını kaydettiği günlükten **günlük maliyeti**, model ve özellik
kırılımını, **kullanıcı başına maliyeti**, **kullanım limiti aşımlarını** ve **ay sonu tahminini** çıkarır.
İnternete bağlanmaz; sağlayıcı hesabınıza erişmez.

## Neden uygulamanın kendi günlüğü?

Sağlayıcıların kullanım panoları toplam maliyeti verir, ama hangi **kullanıcının** ve hangi **özelliğin** ne kadar
harcadığını bilemez. Bunun için uygulamanız her çağrıda kullanıcı kimliğini (takma ad yeterli), özelliği ve
yanıttaki token sayılarını bir satır olarak yazmalıdır. Sağlayıcının CSV dökümü de okunur (zaman, model ve token
sütunları varsa), ama o zaman kullanıcı kırılımı çıkmaz.

| Sütun | Zorunlu | Açıklama |
|---|---|---|
| Zaman | evet | `2026-09-01 10:20:00` ya da ISO biçimi |
| Model | evet | Sağlayıcının model adı |
| Girdi Token / Çıktı Token | evet | Yanıttaki kullanım bilgisi (`input_tokens` / `output_tokens`, `prompt_tokens` / `completion_tokens` adları da tanınır) |
| Önbellek Okuma / Yazma Token | hayır | Önbellek (prompt caching) kullanıyorsanız |
| Kullanıcı, Özellik | hayır | Kullanıcı başına ve özellik başına kırılım için |
| Maliyet | hayır | Varsa fiyat dosyası yerine bu kullanılır |
| Durum | hayır | `hata`, `error`, `failed` gibi değerler hatalı istek sayılır |

## Nasıl hesaplar?

- **Çağrı maliyeti** = (girdi × girdi fiyatı + çıktı × çıktı fiyatı + önbellek okuma × okuma fiyatı + önbellek yazma ×
  yazma fiyatı) ÷ birim (varsayılan 1.000.000 token).
- **Önbellek ve girdi:** Bazı sağlayıcılar girdi token sayısına önbellekten okunanları **dahil eder**, bazıları ayrı
  verir. Sizinki dahil ediyorsa `--girdi-onbellek-dahil` verin; önbellekten okunanlar girdiden düşülür, iki kez sayılmaz.
- **Model eşleştirme:** Fiyat dosyasında birebir ad, yoksa adın başıyla eşleşen en uzun anahtar (`model-a` anahtarı
  `model-a-2026-09` adını da karşılar). Fiyatı bulunamayan modelin maliyeti **0 sayılır** ve "Yüksek" uyarı verilir.
- **Kullanıcı başına** = maliyet ÷ yapay zekâyı kullanan farklı kullanıcı. Uygulamanın tüm aktif kullanıcılarına
  bölünmüş hâli **Uygulama Yönetim Özeti** paketinde, kullanıcı analiziyle birleştirilerek verilir.
- **Ay sonu tahmini** = ayın verideki son gününe kadarki harcama ÷ geçen gün × ayın gün sayısı.
- **Sıçrama** = son gün maliyeti ≥ önceki 7 günün ortalamasının 2 katı (çağrısız gün sıfır sayılır).

## Fiyat ve limit dosyaları

`ornek_veri/fiyatlar.json` **örnek** fiyatlar içerir ve model adları uydurmadır (`model-a`, `model-b`). Kullandığınız
modellerin güncel fiyatlarını sağlayıcının fiyat sayfasından yazın; fiyatlar zamanla değişir.

```json
{ "para_birimi": "USD", "birim": 1000000,
  "modeller": { "model-adi": { "girdi": 3.0, "cikti": 15.0, "onbellek_okuma": 0.3, "onbellek_yazma": 3.75 } } }
```

`limitler.json` (hepsi isteğe bağlı): `aylik_butce`, `gunluk_butce`, `kullanici_gunluk_istek`,
`kullanici_gunluk_maliyet`.

## Kullanım

```bash
python main.py                                         # örnek günlükle dener (kurgusal, Eylül 2026)
python main.py --girdi api_cagrilari.csv --fiyatlar fiyatlar.json --limitler limitler.json
python main.py --girdi gunlukler/ --fiyatlar fiyatlar.json --kur 41,20      # TL sütunları da dolsun
```

Terminal kullanmıyorsanız klasördeki `BASLAT.bat` (Windows) ya da `BASLAT.command` (macOS) dosyasına çift tıklayın.

## Çıktı

- `cikti/yapay_zeka_maliyet.xlsx`: Özet, Günlük (günlük bütçeyi aşan gün kırmızı), Modeller, Özellikler, Kullanıcılar,
  Limit Aşımları, Uyarılar.
- `cikti/yapay-zeka-api-maliyet-raporu.ozet.json`: **Uygulama Yönetim Özeti** paketinin okuduğu günlük maliyet, istek
  ve kullanıcı sayıları.

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

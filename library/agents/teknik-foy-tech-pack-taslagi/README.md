# Teknik Föy (Tech Pack) Taslağı · AI Agent

> Tekstil ve Hazır Giyim › Tasarım ve Koleksiyon › Moda Tasarımcısı · Workers / Workless

Model bilgisi, tasarım notu, ölçü tablosu ve malzeme listesinden **teknik föy (tech pack) taslağı** hazırlar.
- **Kod:** Ölçü tablosunu ve malzeme listesini föye **aynen** aktarır ve hatalarını yakalar.
- **Model:** Ürün tanımını, yapım detaylarını ve talimatları taslak olarak yazar.
- **Onay:** Föy, modelist ve kalite onayıyla kesinleşir.

## Nasıl çalışır?

1. **Dosyaları tanır (kod):** Model klasöründe şunlar beklenir:
   - Model bilgisi: adında "model" geçen `.txt`, `.docx` veya `.pdf`.
   - Ölçü tablosu.
   - Malzeme listesi (BOM).
   - Çizimler ve görseller okunmaz; Kapak sayfasında referans olarak listelenir.
2. **Kontrol eder (kod):**
   - **Ölçü tablosu**
     - Model bilgisindeki beden aralığıyla aynı bedenleri içeriyor mu, ana (numune) beden tabloda var mı?
     - Bir bedenin ölçüsü bir önceki bedenden küçük mü (yazım hatası)?
     - Bedenler arası artışlarda düzensizlik var mı (medyanın 2,5 katından büyük adım)?
     - Eksik tolerans, eksik ölçü, toleranstan küçük beden artışı.
   - **Malzeme listesi**
     - Kumaşta kompozisyon ve gramaj var mı; kompozisyon toplamı %100 mü?
     - Lif adları doğru yazılmış mı? Etikette ticari marka değil genel lif adı kullanılmalı (AB 1007/2011): likra
       veya Lycra → elastan, Tencel → lyocell, naylon → poliamid, "bambu" tek başına lif adı değildir.
     - Zorunlu kalemler eksik mi: dikiş ipliği, marka etiketi, beden etiketi, yıkama/kompozisyon etiketi, poşet.
     - Kumaş renkleri model renk listesinde var mı?
3. **Maskeler:** `--gizle` ile verilen adlar (müşteri markası, tedarikçi) ile telefon ve e-posta maskelenir. Veri
   gönderilmeden önce onayınız alınır.
4. **Yazar (model):** Model şu bölümleri yazar: ürün tanımı, bölüm bölüm yapım detayları, dikiş talimatları,
   baskı/nakış, etiket yerleşimi, ütü ve paketleme, kalite notları, bakım talimatı önerisi ve açık sorular.
   Metindeki girdilerde olmayan cm, mm, °C ve % değerleri **işaretlenir**; bu kontrol uydurma dikiş sıklığı veya
   reçme genişliği gibi değerleri yakalamak içindir.
5. **Rapor:** Excel teknik föy şu sayfalardan oluşur:
   - `Kapak`: model bilgisi ve Tasarımcı / Modelist / Kalite / Müşteri onay alanları.
   - `Ölçü Tablosu`: ana beden kalın.
   - `Malzeme Listesi`.
   - `Yapım Detayları` ve `Talimatlar`: taslak hücreler renkli.
   - `Kontroller`: Onay / Düzeltme sütunuyla.
   - `Revizyonlar`.

   Ayrıca `.md` özet üretilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                        # örnek: TS-2027-014 kadın oversize tişört
python agent.py --girdi ./TS-2027-014 --gizle "Müşteri Markası" "Tedarikçi Adı"
```

| Dosya | Biçim |
|---|---|
| `model_bilgisi.txt` | Satır başına bir alan: `Model Kodu:`, `Model Adı:`, `Sezon:`, `Müşteri:`, `Beden Aralığı: XS-XL` (veya `36-44`, `S, M, L`), `Ana Beden:`, `Renkler:`. Ardından serbest tasarım notu. |
| `olcu_tablosu.csv` | Kod, Ölçü Noktası, Tolerans, beden sütunları (S, M, L… veya 36, 38…) — "Ölçü Tablosu (Beden Serisi Grading)" paketinin çıktısı kullanılabilir. |
| `malzeme_listesi.csv` | Tür (Kumaş / Aksesuar / Etiket / Baskı / Ambalaj), Malzeme Kodu, Malzeme, Kompozisyon, Gramaj, En, Renk, Tedarikçi, Kullanım Yeri, Tüketim, Birim |

## Çıktı

`teknik_foy.xlsx` · `teknik_foy.md`

## Dikkat

- **Taslak niteliği:** Yapım detayları ve talimatlar taslaktır. Dikiş tipleri, sıklıkları ve paylar atölye
  standardınıza göre modelist tarafından tamamlanmalıdır.
- **Bakım talimatı:** Bakım talimatı ve sembolleri (ISO 3758) kumaş testleri (çekme, renk haslığı) yapılmadan
  etikete basılmamalıdır.
- **Etiket kuralları:** Lif adı ve etiket kuralları satılacak pazara göre değişebilir. Türkiye ve AB dışındaki
  pazarlar için müşterinin etiket kılavuzunu esas alın.

## Testler

Testler gerçek API çağırmaz.
- **Kod kontrolleri:** Örnek modelde bilerek bırakılmış beş bulgu kontrol edilir: XS ölçüsü yok, L boyu M'den küçük,
  yaka toleransı yok, "likra" yazımı, beden etiketi eksik.
- **Ek kontroller:** Kompozisyon toplamı ve düzensiz beden artışı test edilir.
- **Model metni:** Uydurma ölçünün yakalanması test edilir.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin (müşteri tasarımları gizlilik
sözleşmesine tabi olabilir) uygunluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.

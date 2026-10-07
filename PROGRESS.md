# İlerleme Günlüğü (PROGRESS)

> Bu dosya, çalışmaya kaldığı yerden devam edebilmek için tutulur (başka bir Claude oturumu dahil).
> Plan ve sıradaki işler: [PLAN.md](PLAN.md). Paket standardı: [STANDART.md](STANDART.md).

## Devam etmek için hızlı bilgi

- Proje kökü: `C:\Users\murat\OneDrive\Desktop\workers-workless`
- Python ortamı: `.venv` (Python 3.14). Testler: `.venv/Scripts/python scripts/kontrol.py --test`
  (Windows konsolunda `PYTHONIOENCODING=utf-8` ile çalıştırın)
- Site: `npm start` → http://localhost:4321 (değişiklikte otomatik yeniden derler)
- Derleme: `node scripts/build.mjs` → `dist/`
- Agent ortak dosyaları: `library/_ortak/llm.py` tek kaynaktır; değiştirince
  `python scripts/kontrol.py --senkronla` çalıştırın.
- GitHub: hesap `muratceme` (gh CLI ile giriş yapıldı, `C:\Program Files\GitHub CLI\gh.exe`).
  Git kimliği depo içinde ayarlı (muratceme / 56095266+muratceme@users.noreply.github.com).

## Günlük

### 2026-10-07 — Oturum 1
- Site tasarımı (Apple tarzı, mod renkleri, komut paleti, terminal demosu).
- Depo yapısı: `catalog/`, `library/`, `site/`, `scripts/`, CI (`.github/workflows/yayin.yml`).
- 4 paket: `cv-raporlama` ve `banka-mutabakati` (kod bloğu + agent), testleri geçiyor.
- İndirme onay penceresi, MIT + SORUMLULUK_REDDI.md, deterministik ZIP + SHA-256.
- İlk commit yapıldı (f2efc30).

### 2026-10-07 — Oturum 2 (devam ediyor)
- GitHub: ilk denemelerde depo oluşturma HTTP 500 verdi, sonra başarılı oldu.
  **Depo:** https://github.com/muratceme/workers-workless (public) · **Site:** https://muratceme.github.io/workers-workless/
  Pages "GitHub Actions" kaynağıyla açık; `gorev-talebi` etiketi oluşturuldu. CI yeşil (Windows+Linux, Py 3.10/3.13).
- Araştırma: İş Bankası/Ziraat organizasyon şemaları, Anadolu/Ana Sigorta, gerçek ilan başlıkları
  (kariyer.net indeksi, eleman.net, kampusum, secretcv, yenibiris). kariyer.net doğrudan erişimi bot korumasıyla
  engelliyor (CAPTCHA aşılmadı); arama motoru + toplayıcı siteler kullanıldı. Kaynaklar: catalog/KAYNAKLAR.md
- **Katalog yeniden kuruldu:** catalog/ortak.mjs (13 ortak departman) + catalog/sektorler/*.mjs
  → 16 sektör, 94 departman, 176 rol, 430 tekil görev. Görev kayıt defteri: aynı görev başka rolde adıyla anılır
  (build.mjs farklı tanımlı aynı adı reddeder).
- server.js: derleme artık ayrı süreçte (alt modül önbellek hatası düzeltildi).
- **Yeni paketler (hepsi kod bloğu, testli, CI yeşil):**
  - `vkn-tckn-iban-dogrulama` — algoritmalar python-stdnum (GitHub) kaynak koduyla ve bilinen örneklerle doğrulandı
  - `e-fatura-okuma-ve-listeleme` — UBL-TR; GİB resmî tevkifat örneğiyle test (tests/fixtures, MIT kaynak belirtildi)
  - `brut-net-maas-hesaplama` — 2026 parametreleri; 40.207,53 / 28.075,50 / 57.881,23 gibi bağımsız değerlerle birebir
  - `kidem-ve-ihbar-tazminati-hesaplama` — tavan dönemleri 2025-2026, bağımsız örnekle doğrulandı
  - `yillik-izin-hakedisi-hesaplama` — 4857 md. 53
- **Ortak çekirdekler:** `library/_ortak/bordro.py`, `library/_ortak/tr_parametreler.json` (yıl bazlı, kaynaklı).
  `ortak_dosyalar` artık kod bloklarında da kullanılıyor (ör. e-fatura → vkn modülü).
- **Bulunan/düzeltilen genel hatalar:** openpyxl read_only dosya kilidi (Windows); Türkçe "İ".lower() → "i̇" başlık
  eşleşme hatası (her yeni pakette `kucuk()` yardımcı fonksiyonu kullanın); Windows cp1254 konsol ("[OK]" kullanın).
- Yeni paketler: `risk-degerlendirmesi-fine-kinney`, `aql-orneklem-plani` (MIL-STD-105E/ISO 2859-1; GitHub'daki
  "doğrulanmış" bir veri setinde hatalı Seviye I sütunu bulundu, kullanılmadı), `spc-kontrol-grafigi-ve-cp-cpk`,
  `oee-hesaplama` (klasik oee.com örneğiyle test), `temsilci-ihtiyaci-hesaplama-erlang-c` (yayımlanmış örnekle test).
- **Mevzuat dersi:** Form Ba-Bs 565 Sıra No'lu VUK Tebliği (RG 25.09.2024) ile 01.10.2024'ten itibaren kaldırılmış.
  Katalogdan çıkarıldı, yerine "Gelen e-Fatura Kayıt Kontrolü" eklendi. **Mevzuata bağlı her görevi yazmadan önce
  yürürlükte olduğunu doğrula.**
- kontrol.py artık her paketin `--help` komutunu da çalıştırıyor (argparse "%" hatası böyle yakalandı).
- CI: art arda push'larda Pages yayını çakışıyordu → workflow'a `concurrency` eklendi.
- Yeni paketler: `gelen-e-fatura-kayit-kontrolu`, `alacak-yaslandirma` (FIFO), `nakit-akis-tahmini` (13 hafta).
  **Toplam 17 paket** (13 kod bloğu yeni + 4 ilk paket), hepsi testli.

### 2026-10-07 (devam oturumu)
- Yeni paketler: `cari-hesap-mutabakati` (karşı taraf yönü otomatik, KDV/tevkifat farkı tahmini, yoldaki kalemler,
  mutabakat mektubu), `butce-gerceklesen-sapma-raporu` (önek eşleşmesi 770.02.001→770.02, Tekdüzen gelir/gider yönü,
  önemlilik eşiği, zamanlama uyarısı, yıl sonu tahmini), `hakedis-hesaplama` (inşaat; kümülatif icmal).
- Hakediş için doğrulanan mevzuat: yapım işlerinde KDV tevkifatı 4/10 — belirlenmiş alıcı veya KDV dahil bedel
  ≥ 5.000.000 TL (KDV GUT I/C-2.1.3.2.1, Seri 35, 01.03.2021); yıllara yaygın inşaat stopajı %5 (GVK 94/3, 3491 s. CBK,
  01.03.2021), demiryolu/gemi/nükleer %1. Kaynaklar catalog/KAYNAKLAR.md'de.
- **Toplam 20 paket**, kontrol 20/20.

- Devam: `pazaryeri-siparis-karlilik-hesabi` (Trendyol resmî Satıcı Bilgi Merkezi: komisyon KDV dahil rakamdan
  hesaplanır, ekstra KDV yok — blogların "+%20 KDV" hesabı yanlış; e-ticaret stopajı %1, 9284 s. CBK),
  `teklif-karsilastirma`, `mali-tablo-rasyo-analizi` (mizan hiyerarşisi, ortalama bakiye).
- Hata düzeltmesi: slug fonksiyonu şapkalı harfleri (â, î, û) tanımıyordu ("Kârlılık" → "k-rl-l-k"); build.mjs ve
  app.js düzeltildi (arama normalleştirmesi dahil).
- Agent'lar (hepsi sahte model yanıtıyla testli; gerçek API ile henüz test EDİLMEDİ):
  `sozlesme-on-inceleme` (maddeleme, 20 konuluk kontrol listesi, prompt'ta yalnız doğrulanmış kanun referansları),
  `musteri-talebi-siniflandirma-ve-cevap-taslagi` (bilgi bankasına dayalı, kod kuralları modelin üstünde, KVKK 30 gün),
  `urun-aciklamasi-uretme` (üret → kodla denetle → düzelt döngüsü), `hasar-dosyasi-ozeti` (belgeler arası tutarlılık,
  plakalar takma adla maskelenir). Yeni ortak modül: `library/_ortak/belge.py` (PDF/DOCX/TXT).
- **Toplam 27 paket** (21 kod bloğu + 6 agent), kontrol 27/27, CI yeşil.

### KALDIĞIM YER
PLAN.md Aşama 2'deki sıra: stok ABC/XYZ + yeniden sipariş noktası → otel doluluk/ADR/RevPAR → tekstil 4 puan kumaş
kontrolü → ölçü tablosu grading → agent'lar (toplantı notu→aksiyon, müşteri yorum analizi, iş ilanı+mülakat, 8D,
mali tablo yorum raporu, PDF fatura okuma). Klasör adı = katalogdaki görev adının slug'ı (build.mjs'deki slug).
Komut zinciri: `PYTHONIOENCODING=utf-8 .venv/Scripts/python scripts/kontrol.py --senkronla` → `--test` →
`node scripts/build.mjs` (hata verirse çıkış kodu 1; `| tail` ile gizlemeyin) → commit + push.

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

- Yeni kod blokları: `stok-abc-xyz-analizi`, `yeniden-siparis-noktasi-hesabi` (ders kitabı örnekleriyle test),
  `doluluk-adr-ve-revpar-raporu` (GY 364 gün hizalı), `kumas-kontrol-4-puan-sistemi-raporu` (ASTM D5430 uygulaması;
  yayımlanmış 14,66 örneğiyle test), `olcu-tablosu-beden-serisi-grading`.
- Yeni agent'lar: `8d-rapor-taslagi`, `misafir-yorum-analizi`, `is-ilani-metni-hazirlama`, `mulakat-soru-seti-hazirlama`.
- Yeni ortak modül: `library/_ortak/ayrimcilik.py` (4857 md.5, 6701 md.3/6/7, 6356 md.25, KVKK md.4 dayanaklı kural
  tabanlı tarayıcı; yanlış alarm testleri var). `llm.maskele` artık sabit hat/0850 numaralarını da maskeliyor
  (başında 0 veya +90 şart, sipariş no gibi 10 haneli sayılar etkilenmez).
- Not: Bu bilgisayarda Excel/LibreOffice yok; mülakat formundaki Excel formülleri (AVERAGEIF, SUMPRODUCT) hesaplatılarak
  doğrulanmadı, yalnız yapısı test edildi.
- **Toplam 36 paket** (26 kod bloğu + 10 agent), kontrol 36/36.

- Yeni paketler: `mulakat-degerlendirme-ozeti`, `gorusme-kaydi-ozeti`, `destek-talebi-siniflandirma` (agent);
  `haftalik-uretim-cizelgesi`, `malzeme-ihtiyac-planlamasi-mrp`, `cv-on-eleme`, `yonetim-raporu-kpi`,
  `banka-hareketlerinden-muhasebe-fisi-onerisi` (kod).
- Hata: Türkçe küçük harf dönüşümü "IBAN"/"ID" başlıklarını "ıban"/"ıd" yapıyordu; ilgili paketlerde düzeltildi.
- Depoda `core.autocrlf=false` yapıldı, çalışma kopyası LF'ye normalleştirildi (CRLF uyarıları bitti).
- **Toplam 44 paket**, kontrol 44/44.

### 2026-10-08
- Yeni kod blokları: `stok-sayim-fark-analizi` (bekleyen hareket düzeltmesi, lokasyonlar arası mahsup, kural tabanlı
  neden ipuçları: fazla sıfır, rakam yer değiştirme, koli/adet, varyant karışıklığı; ikinci sayım listesi; 197/397),
  `mizan-kontrolu` (TDHP doğa tablosu, alt hesap ters bakiyeleri ve virman önerileri, kasa/131/KDV/7A yansıtma,
  önceki aya göre değişim ve kümülatif azalma; yıl başı otomatik algılanır), `puantaj-kontrolu` (İş K. md. 41/46/47/56/63,
  2025-2027 bayramları, SGK prim günü ve eksik gün nedeni), `urun-maliyeti-hesaplama` (çok seviyeli reçete, rota,
  iş merkezi ücretleri, döviz kuru, brüt kâr marjı).
- Bu paketlerde başlık eşleştirme `katla()` ile yapılıyor (ı/i ve Türkçe karakter farkı yok sayılır). Eski paketlere
  yaymak hâlâ açık iş.
- **Toplam 48 paket**, kontrol 48/48.
- Devam: `sik-sorulan-sorulara-cevap-taslagi` (agent; BM25 + Türkçe 5 harf kök, yalnız eşleşen bölümler gönderilir,
  uydurma kaynak/parola isteyen taslak kodla yakalanır, bilgi bankası boşlukları), `supheli-islem-senaryo-taramasi`
  (9 senaryo, senaryolar.json — yasal eşik koda gömülmedi; 5549 md. 4/2 ifşa yasağı uyarısı), `mizandan-mali-tablo-hazirlama`
  (MSUGT bilanço/gelir tablosu, ters bakiye virmanları; mizan-kontrolu çekirdeğini ortak dosya olarak kullanır, çekirdeğe
  TDHP_AD eklendi), `kumas-ve-aksesuar-ihtiyac-hesabi`, `ibnr-rezerv-tahmini-zincirleme-merdiven` (Taylor-Ashe ile
  18.680.856 / Mack 2.447.095 birebir), `fazla-mesai-hesaplama` (puantaj çekirdeğini kullanır; UBGT saatlerinin 45 saate
  dahil edilmesi tartışmalı → --ubgt-haric seçeneği), `ilk-dagitim-alokasyon-plani`.
- **Toplam 55 paket**, kontrol 55/55.
- Devam: `siparis-termin-plani-t-a` (geriye planlama + ileri tahmin; puantaj takvim çekirdeği), `urun-sorusu-cevap-taslagi`
  (agent; SSS arama çekirdeğini kullanır; iletişim bilgisi/uydurma sayı/teslim sözü kodla yakalanır),
  `performans-degerlendirme-ozeti` (agent; puanlar koddan, takma ad P1/Y1, değerlendirici eğilimi; ortak ayrımcılık
  tarayıcısına değerlendirme yorumu desenleri eklendi: "yaşı gereği", "çocuklu olduğu için", "sağlık sorunları nedeniyle",
  "doğum iznine çıkacağı" — iş ilanı yanlış alarmları test edildi), `sinav-sonuc-ve-madde-analizi` (p, d, r, KR-20 elle
  hesapla test), `zafiyet-tarama-onceliklendirme` (KEV + EPSS + CVSS + varlık; örnek CVE'ler kurgusal CVE-2099-…),
  `yaptirim-listesi-taramasi` (BM XML yapısı resmî dosyadan doğrulandı; Türkçe/transliterasyon normalleştirme).
- **Ders:** test ve commit'i `&&` ile zincirle (bir kez `;` yüzünden kırık commit push edildi, hemen düzeltildi).
  Python heredoc'ta `` yazınca dosyaya backspace () düşebiliyor — regex düzenlemelerini Edit aracıyla yap.
- **Toplam 61 paket**, kontrol 61/61.

- Devam: `kesim-kat-plani` (açgözlü + kapanışta 1-2 pastal, min kat), `koli-listesi-packing-list-hazirlama`,
  `hasar-siklik-ve-siddet-analizi` (değerleme tarihi; güvenilirlik 1082), `satis-hedef-gerceklesme-raporu`,
  `reklam-kampanyasi-performans-raporu` (TR/EN başlık eşanlamlıları, 2.500 binlik ayırıcı), `otomatik-sevkiyat-replenishment-onerisi`
  (depo yetersizse en düşük stok günlü mağazaya önce), `taseron-hakedis-kontrolu` (sözleşme fiyatı / saha metrajı / onaylı
  önceki; bu dönem onayı talep edileni aşamaz; vergi ve kesintiler Hakediş Hesaplama paketinde).
- **Toplam 68 paket**, kontrol 68/68.

- Devam (2026-10-08): `erken-uyari-sinyalleri-listesi` (gecikme/limit/KKB çek-senet/haciz/ciro düşüşü puanlaması; 30/90 gün
  Aşama 2/3 notu; puanlar ayarlar.json'da, örnek), `gece-denetimi-night-audit-kontrolu` (açık folyo, skip/sleep, fiyat kodu,
  kasa mutabakatı, oda geliri/ADR/RevPAR), `konsolide-finansal-rapor` (mizan + mali tablo çekirdeklerini ortak dosya olarak
  kullanır; cari/gelir-gider/stoktaki kâr/sermaye eliminasyonu, şerefiye 261, KGO özkaynakta "5KG" ve kârda "59K" satırı;
  elle hesap 9.470.000 / 1.030.000 birebir), `kredi-teklif-dosyasi-hazirlama` (agent; rasyolar + KKB + katsayılı teminat
  karşılaması koddan; unvan/VKN/ortak adları maskeli; uydurma sayı yakalanır; bankacılık sırrı uyarısı → Ollama önerisi),
  `teknik-foy-tech-pack-taslagi` (agent; ölçü sırası/düzensiz artış, kompozisyon %100, likra→elastan lif adı, zorunlu
  etiketler; ölçü/BOM koddan aynen, birimli uydurma değerler işaretlenir).
- **Ders:** mizan çekirdeğinin para()'sı "750.000"ı 750 okur — tek noktalı binlik yazım için paket içinde sarmalayıcı
  yazıldı (çekirdeğe dokunulmadı; ileride çekirdekte düzeltilebilir).
- **Toplam 73 paket**, kontrol 73/73.

- Devam (2026-10-08, 2. tur): `sikayet-analizi` (agent; konu listesi verilmezse model örneklemle 5-12 konu önerir,
  sonra 25'lik paketlerle sınıflar; eğilim/Pareto/hedef süre/30 gün tekrar koddan; listede olmayan konu → "Diğer"),
  `ihale-dokumani-ozeti` (agent; her kalem kaynak madde + birebir alıntı, kod alıntıyı belgede arar; işin süresi ve günlük
  gecikme cezası belgeler arası çelişki taraması; özette atlanan konu taraması; --teklif ile teminat/iş deneyimi/ciro/ceza
  tutarları; örnek ihale kurgusal, 300↔330 gün çelişkisi bilerek), `siparis-termin-takibi` (teyit termini esas; 1-7/8-30/30+
  öncelik; geçmiş teslimlerden ort. gecikme → tahmini teslim; ihtiyaç tarihi riski; tedarikçi hatırlatma metni taslağı),
  `musteri-talimat-kontrolu` (IBAN mod 97 + TR rezerv hane; bitişik Türkçe yazıyla tutar ayrıştırıcı; havale/virman/EFT tür
  kuralları, EFT yalnız TL; münferit/müşterek imza + limit + süresi dolmuş yetki; mükerrer), `toplu-urun-yukleme-sablonu-hazirlama`
  (belirli pazaryeri taklit edilmez: JSON eşleştirme + kullanıcının indirdiği şablon; GTIN kontrol hanesi, varyant tekrarı,
  KDV 18 → izinsiz; .xlsx şablon kopyalanıp doldurulur).
- **Ders:** `;` ayırıcılı CSV'de çoklu değer (görsel listesi, imzalayanlar) için `|` / `,` kullan — aynı ayırıcı sütunu böler.
  Windows'ta Python write_text CRLF yazar (git normalize ediyor, sorun değil).
- **Toplam 78 paket**, kontrol 78/78.

- Devam (2026-10-08, 3. tur): `log-anomali-tespiti` (access/auth/uygulama logu satırdan tanınır; SSH/web parola
  deneme + ardından başarılı giriş = kritik, 404 taraması, saldırı imzaları, medyan+4·MAD ani artış, normalleştirilmiş hata
  imzası; kontrol.py urllib'i yasaklıyor → URL çözme elle), `resmi-yazi-haciz-muzekkere-takibi` (yalnız İİK 89/1=7 gün,
  89/2=15 gün kodda; diğer türlerde süre yazıdan; son gün tatile denk gelirse ilk iş günü; 89/1 cevapsızken 89/2 → kritik),
  `geri-odeme-kapasitesi-analizi` (CFADS = FAVÖK − vergi − yatırım − NİS artışı; aylık ödeme planlarından DSCR; 6 senaryo;
  ikili aramayla azami kredi ve kırılma satış düşüşü), `portfoy-musteri-firsat-listesi` (kural tabanlı; gecikme/limit
  doluluğunda kart-kredi önerilmez, sigorta krediye bağlanmaz notu, İYS izni yoksa yalnız şube), `firma-istihbarat-raporu-derleme`
  (agent; sicil ayrıştırma, ortaklık toplamı, KKB ↔ istihbarat farkı, istihbaratı alınmamış banka, olumsuz kayıt yaşı;
  sicildeki ortak/yöneticiler otomatik [KİŞİ-n]/[ŞİRKET-n]).
- **Ders:** Git Bash heredoc'ta bazı tırnak kombinasyonları "unexpected EOF" veriyor — uzun dosyaları Write ile yaz.
  `node scripts/build.mjs | tail -1` hata kodunu yutuyor; bir kez OneDrive kilidi yüzünden ENOENT verdi, yeniden çalıştırınca geçti.
- **Toplam 83 paket**, kontrol 83/83.

- Devam (2026-10-08, 4. tur): `kredi-komitesi-degerlendirme-notu` (agent; limit/teminat/rasyo tabloları + metinler;
  mevcut ve tesis edilecek teminat karşılaması ayrı, rasyo yönü addan, DSCR<1 / 1–1,2, isteğe bağlı BK md. 54 %25 sınırı;
  model sonrası: yüksek bulgu varken "Olumlu" ve şartlarda olmayan tesis edilecek teminat uyarısı), `sube-teftis-veri-analizi`
  (personelin kendi/yakın hesabı, mesai dışı/tatil, limit aşımı onaysız, kendi onayı, onaylayan limiti, sık iptal, iptal sonrası
  farklı tutar, bölünmüş nakit (eşik parametre, MASAK eşiği yazılmadı), yoğunlaşma; örnekte bilerek 12 bulgu),
  `teftis-raporu-bulgu-taslagi` (agent; dayanak yalnız kullanıcının verdiği mevzuat/yönerge metninden, maddelere bölünür,
  alıntı birebir doğrulanır; örnek yönerge kurgusal), `risk-degerlendirme-ozeti` (agent, sigorta; 25 konuluk kontrol
  listesi, koruma önlemi + olumsuzlama = olumsuz gözlem, emtia bedeli < maks. stok uyarısı, R-numaralı öneri kapsamı),
  `hasar-dosyasi-evrak-eksik-kontrolu` (evrak_listesi.csv örnek ve düzenlenebilir; koşullu evraklar; Türkçe ek toleranslı
  kelime-başı eşleştirme; iç takip talep yazısına girmez).
- **Ders:** Anahtar kelime eşleştirmede kısa kökler tuzak: "foto" → "fotokopisi". Kelime başı + uzun kök kullan.
  Rastgele örnek veride planlı anomalilerin müşteri numaralarını arka plandan dışla (yoksa istemeden ek bulgu çıkar).
- **Toplam 88 paket**, kontrol 88/88.

### KALDIĞIM YER
Sıradaki (paketsiz, puan 4): Teminat Kapsam Kontrolü (agent, sigorta), ardından listeyi yeniden üret ve puan 4'leri sürdür…
Sağlık/SGK faturası ve epikriz sona bırakıldı. Agent'lar gerçek API ile henüz denenmedi (kullanıcı anahtar verecek).
Listeyi üretmek için: geçici bir node betiğiyle catalog/catalog.mjs DEPTS'ten görevleri slug'layıp library/ ile karşılaştır.
İyileştirme notları: başlık eşleştirmede ı/i ve Türkçe karakter katlayan ortak bir normalizasyon (tüm paketlerde);
mizan çekirdeği para() "750.000" binlik yazımı.
Komutlar: `kontrol.py --senkronla` → `--test` → `node scripts/build.mjs` → commit + push.

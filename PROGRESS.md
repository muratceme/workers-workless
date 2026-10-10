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

- Devam (2026-10-08, 5. tur): `teminat-kapsam-kontrolu` (agent; poliçe süresi / prim / teminat var mı, eksik sigorta oranı,
  "%2, en az 5.000 TL" muafiyet ayrıştırma, limit → ön hesap; model kapsam dışı kalemi tutar + alıntıyla verirse ikinci
  ön hesap; örnek: dahili su, korozyon istisnası belirsiz, zemin istifi), `reasurans-hesap-cetveli-hazirlama` (kotpar /
  eksedan oranı, kapasite aşımı → fakültatif, iptal ilk poliçe oranıyla, prim depo + faiz, bakiye ve kuruş farkı son
  reasüröre, cash call, kâr komisyonu yalnız bilgi), `coklu-sirket-teklif-karsilastirmasi` (talep listesiyle kelime
  benzerliği eşleştirmesi, muafiyet istenen limitte TL'ye çevrilip kıyaslanır, süre muafiyeti kıyaslanmaz, süresi dolmuş
  teklif uygun değil; müşteriye Markdown tablo), `gumruk-evrak-tutarlilik-kontrolu` (belge_bilgileri alan × belge
  matrisi, sayısal alanlarda çoğunluk değeri referans + tolerans, unvan eki farkı yok sayılır, kalem miktar / GTİP ilk 6 /
  menşe, FOB-FAS-CFR-CIF + CMR uyarısı, GTİP 12 hane bilgi), `haccp-kritik-kontrol-noktasi-izleme-analizi` (kritik /
  operasyonel limit, kategorik metal dedektör, DF kaydı yok, izleme boşluğu, eğilim, tekrar eden değer; limit çizgili grafik).
- **Ders:** Örnek CSV metin alanlarında `;` kullanma (DF açıklamaları sütun kaydırdı). Türkçe ek gerektiren dinamik
  metinlerden kaçın ("6'i", "08:15'den") — ekten bağımsız kalıp kullan ("6 iptal", "08:15 itibarıyla").
- **Toplam 93 paket**, kontrol 93/93.

- Devam (2026-10-08, 6. tur): `kestirimci-bakim-uyarilari` (ISO 10816-3 grup / temel → B/C uyarı, C/D alarm; nominal
  akım alarm, ×0,90 uyarı; 30 günlük doğrusal eğilim + R² ≥ 0,6 ile alarma kalan gün; Kriter II ani değişim; çözünürlüğe
  duyarlı sabit değer: medyan |Δ| ≥ 2 × çözünürlük; ölçüm gecikmiş / yok; çoklu belirti), `numune-takip-cizelgesi` (proto /
  fit / size set / PP / SMS / TOP, termin = istenen veya talep + hazırlık; yanıt gecikmesi, açılmamış revizyon, fit onayı
  olmadan PP, kesim öncesi PP onayı; model × tür matrisi, zamanında gönderim performansı), `urun-maliyet-hesaplama-costing`
  (kumaş tüketim × fire × fiyat, CM = SAM × dk maliyeti ÷ verim, sabit ÷ adet, GG, FOB = maliyet ÷ (1 − marj − komisyon);
  kur çevirisi, hedef fiyatta marj ve gereken düşüş, kur ±%5 / kumaş ±%10 duyarlılığı), `uretim-hatti-yukleme-plani`
  (puantaj_cekirdek tatilleri; arife yarım gün; EDD sırası + en erken bitiren uygun hat, sabit atama; öğrenme eğrisi
  50/70/85; gün ortasında devralma; tampon 2 iş günü; gereken adet/gün; hat × gün planı), `kumas-siparis-termin-takibi`
  (tip bazında aşama şablonu, son tamamlanan aşamadan tahmin, gecikmiş aşama bugün biter, lab dip onayı yoksa boyama bugün
  başlar; eksik teslim toleransı; kritik kumaş → kesim / sevk kayması; tedarikçi performansı).
- **Ders:** Gürültülü ama düşük çözünürlüklü ölçümlerde "art arda aynı değer" doğaldır; sabit değer kuralını serinin
  değişkenliğine bağla. argparse help metninde `%` → `%%`. Yerelde localhost sunucusu açıkken `build.mjs` dist
  dosyalarına erişemeyebiliyor (ENOENT): sunucuyu durdur → build → yeniden başlat.
- **Toplam 98 paket**, kontrol 98/98.

- Devam (2026-10-08, 7. tur): `fason-atolye-is-takibi` (kısmi dönüşler, açık adet, GECİKTİ / yaklaşıyor, tempo: süre ≥ %50 ve
  dönüş oranı 30 puan geride, hata oranı eşiği ve 2× Yüksek, tamir, "Kapandı" + fire toleransı, fazla dönüş; atölye
  performansı; dönem filtreli hakediş = sağlam × fiyat), `arac-ve-sefer-takip-listesi` (ülke bazında aşama süreleri saat
  olarak, "*" varsayılan; gecikmiş aşama şimdi biter; yükleme gecikmesi; teslim gecikmesi 24 saat eşiği; eski konum; CMR;
  araç çakışması; araç durumu / müsait zaman), `parsiyel-yuk-konsolidasyon-plani` (LDM = ⌈palet ÷ istif⌉ × en × boy ÷ 2,4;
  ödenebilir = max(kg, m³×333, LDM×1.750); bölge sırası en erken son yükleme; ADR yük önce ve yalnız ADR araca, ADR'li araç
  ADR'siz yüke en son; first-fit + bölge sonunda küçültme; düşük doluluk + bekletilebilir önerisi), `dilekce-taslagi`
  (agent; HMK 119/1 kontrol listesi a–h, TCKN / VKN kontrol hanesi, zorunlu arabuluculuk ipucu (6325 md. 18/A), başlık /
  taraflar / deliller / ekler kodla; model açıklama + olay özetinden birebir dayanak alıntısı + delil no, hukuki sebepte
  madde no yalnız verilen mevzuattan, kanun sayısı beyaz listeden; adres ve kimlik no modele gitmez),
  `hasta-iletisim-mesaji-cevirisi` (agent; mesaj başına çağrı; ad parçaları + uluslararası telefon + pasaport maskesi;
  doz / ölçü (sayı + birim, Kiril birimler) çeviride korunmuş mu, kırmızı bayrak ↔ aciliyet, acil cevapta acil servis,
  cevapta kaynaksız sayı, cevap ↔ Türkçe karşılık sayıları, terimce, alfabe ↔ dil).
- **Ders:** `katla("*")` boş dizgi döner — varsayılan anahtar olarak "" ara. Başlıkta "m³" katlanınca "m" kalır; eş
  adlara "hacim m" ekle. Sahte model yanıtını mesaj gövdesine göre seç (terimce gibi ortak bölümler her mesajda var).
- **Toplam 103 paket**, kontrol 103/103.

- Devam (2026-10-08, 8. tur): `provizyon-talebi-degerlendirme` (sıralı karar: anlaşmasız → süre dışı → teminat yok → genel
  şart istisnası (ICD-10 öneki) → poliçedeki ön mevcut durum → teminat / tanı bekleme süresi (ilk giriş tarihinden) →
  inceleme (ICD yok, prim gecikmede, ilk 90 günde kronik tanı) → katılım payı + kalan yıllık limit; onaylar limitten
  düşülür; kurum cevabı taslağı), `saglik-faturasi-kontrolu` (fiyat farkı, hesap hatası, paket içi hizmet, mükerrer,
  provizyon başına adet sınırı, listede yok / provizyon öncesi işlem → inceleme; provizyon ret / inceleme / yok; şirket
  payı provizyon tavanıyla sınırlı; kesinti listesi), `mulakat-takvimi-planlama` (rol paneli, pozisyon yetkisi, günlük
  sınır, tampon, en dar müsaitlik önce, 15 dk adım; yerleşemeyen gerekçesi rol bazında; davet metni + .ics),
  `personel-ozluk-dosyasi-eksik-evrak-takibi` (koşullu evrak listesi Alan=/!=/~, teslim süresi 0 = işe başlamadan önce,
  boş = 30; geçerlilik ay; matris + hatırlatma metinleri), `ucret-bandi-karsilastirmasi` (pozisyon bandı önce, kısmi
  süreli tam zamanlı karşılık, compa-ratio, asgari ücret tr_parametreler.json'dan, sıkışma ≥2 yıl kıdem + performans,
  bütçe etkisi, cinsiyet karşılaştırması yalnız bilgi).
- **Ders:** Sahte / örnek veri üretirken özel durum satır indekslerini üretimden sonra yazdırıp doğrula. Sıralı kural
  motorlarında "ilk uyan kural" sırasını README tablosunda göster.
- **Toplam 108 paket**, kontrol 108/108.

- Devam (2026-10-09, 9. tur, 10 paket): `zam-butcesi-simulasyonu` (Performans × Konum → Oran + Sabit Tutar, '*' ve en özel
  satır; compa-ratio konumu; asgari ücret tabanı — yılı yoksa önceki yıl + uyarı; kıst, bant üstü sınırı → tek seferlik,
  yuvarlama; bütçeye ölçekleme ikiye bölmeyle; işveren maliyeti SGK tavanlı), `performans-primi-hesaplama` (artan / azalan /
  hedefi 0 hedef, hedef tavanı, skala eşik + doğrusal / kademeli + tavan, şirket çarpanı, kıst, ayrılana ödeme politikası,
  bütçe aşımı), `personel-devir-orani-analizi` (ortalama = dönem başı + ay sonları; ayrılış türü nedenden anahtar
  kelimeyle; kıdem bandı paydası; 90 gün erken ayrılma kohortu; 1,5 kat sinyali; bağımsız hesapla doğrulandı),
  `egitim-ihtiyac-analizi` (açık × kritiklik + düşük performans + talep → öncelik; algı farkı; eğitim planı grup /
  bireysel), `calisan-baglilik-anketi-analizi` (**agent**: puanlar kodda, anonimlik birleştirme, eNPS; yorumlar
  maskeli ve birimsiz gönderilir; alıntı doğrulama; başlıklar yorum kimliğine dayanmalı; anahtar kelimeyle hassas yorum
  güvenlik ağı; prompt.md iki bölümlü), `fatura-kdv-tutarlilik-kontrolu` (10.07.2023 oran değişimi, tevkifat oran
  kümesi, yuvarlama sınıfı, e-fatura paketinin Excel çıktısını okur), `donem-sonu-kur-degerleme` (işaretli bakiye,
  646 / 656, TL kalıntısı, avanslar hariç, ters bakiye, kur sapması), `cari-hesap-ekstresi-hazirlama` (devir,
  yürüyen bakiye, FIFO açık kalem, vade yalnız fatura türü kalemlerde, mutabakat metni, ayrı dosyalar),
  `masraf-fisi-kategorileme` (öncelikli kural dosyası, fonksiyon hesabı + alt hesap, belge türüne göre KDV, KKEG,
  ödeme hesapları 309 / 195 / 335), `personel-maliyet-butcesi` (kadro + yan hak kapsamı, SGK tavanı ve teşvik, zam
  mevcut kadroya, gerçekleşme sapması).
- **Ders:** Başlangıç ayı boş = mevcut kadro; "ilk ay" ile karıştırma (Ocak zammı hatası testte yakalandı). Ödeme
  kayıtlarına varsayılan vade yazma (fazla tahsilat "vadesi geçmiş" görünüyordu).
- **Toplam 118 paket**, kontrol 118/118.

- Devam (2026-10-09, 10. tur, 10 paket): `standart-fiili-maliyet-sapma-analizi` (711/712/721/722, GÜG üç sapma 731/732/733;
  toplam = fiili − standart mutabakatı; elle hesaplanmış örnek), `stok-degerleme` (FIFO / hareketli / dönem sonu AO yan yana;
  alış ve satış iadesi; eksi stokta son maliyet + sonraki girişte düzeltme → değer korunumu; NGD testi),
  `kdv-beyannamesi-on-kontrolu` (liste ↔ mizan 190/191/391 ↔ beyanname; kısmi tevkifatta satıcı KDV − tevkif; KDVK 29/3
  indirim süresi; beyanname aritmetiği; sade Alan;Tutar beyan dosyası), `cek-senet-portfoy-takibi` (TTK 796 ibraz süresi,
  banka karşılık açığı kümülatif, karşılıksız keşidecinin diğer evrakı, yoğunlaşma, ciro riski, haftalık vade),
  `yatirim-fizibilitesi-npv-irr` (vergi + 5 yıl zarar mahsubu, reel/nominal dönüşüm, IRR ikiye bölme — Wikipedia örneği
  %5,96 ile test, MIRR, geri dönüş), `senaryo-ve-duyarlilik-analizi` (sürücü tabanlı tek yıllık model, %/puan/= senaryo,
  tornado, başa baş hacim ve kur), `departman-butcelerinin-konsolidasyonu` (klasördeki xlsx şablonları, hücre adresli hatalar,
  Excel TOPLA metin sayıyı atlar uyarısı, hesaplanmamış formül, mükerrer departman), `kullanici-yetki-gozden-gecirme`
  (sicil → e-posta → ad eşleşmesi, ayrılıştan sonra giriş, sistemler arası SoD, rol matrisi '*', yönetici onay listesi),
  `kvkk-kisisel-veri-isleme-envanteri` (md. 5 / md. 6 / md. 9 — 7499 sayılı Kanun 2024; özel nitelikli veri anahtar
  kelime, kısa kelimeler tam kelime), `bt-maliyet-raporu` (yıllık ödemeler aylara yayılır, rapor ayı sonrası peşin;
  tahakkuk + peşin = ödenen; atıl lisans, mükerrer abonelik, bulut artışı, yenileme takvimi).
- **Ders:** Tahakkukta rapor ayından sonraki aylar da "dönem"e yazılıyordu (bütçe ↔ kategori sayfası farkı testte görüldü).
  Decimal ile float sınır karşılaştırması (`< 0.8`) tam sınırda yanlış sonuç verir; Decimal("0.8") kullan.
- **Toplam 128 paket**, kontrol 128/128.

- Devam (2026-10-09, 11. tur, 10 paket): `satin-alma-talebi-konsolidasyonu` (onay durumu, birim eş adları, net ihtiyaç = talep −
  (stok − emniyet) − açık sipariş, asgari sipariş + kat, termin riski, 3 teklif eşiği, tedarikçi × malzeme teklif listesi ve
  ayrı teklif formları), `tedarikci-performans-degerlendirme` (kısmi teslim birikimi, eksik kapama / Satır Durumu, açık gecikmiş,
  kriter ağırlıkları ve verisi olmayan kriterin çıkarılması, çeyreklik trend, karne metni; bir tedarikçi elle doğrulandı),
  `harcama-analizi` (ABC önceki kümülatif paya göre, tedarikçi adı birleştirme, ay içi fiyat farkı, medyana göre fiyat artışı —
  enflasyon, pazarlık öncelik kuralları), `dava-ve-durusma-takvimi` (HMK 92/93/104 + İYUK süre tablosu, adli tatil 7 Eylül,
  bayramlar dosyadan, yarım gün uyarısı, iç hedef, duruşma çakışması, .ics), `icra-takip-durum-raporu` (basit faiz + TBK 100 mahsup
  sırası, İİK 62/67/68/78/149/168 süreleri, işlemsiz dosya, fazla tahsilat), `mevzuat-degisikligi-ozeti` (**agent**: fihrist
  ayrıştırma, başlık sınıflandırma, yürürlük maddesi kodda birebir + tarih kuralları, değişiklik alıntı doğrulama, bülten .md;
  örnek veri tamamen kurgusal), `hizmet-alimi-fatura-kontrolu` (Fiilî + asgari günlük / Puantaj kişi-gün/30 / Sabit, fiyat
  geçerlilik dönemi, KDV oranı, tevkifat kesri, mükerrer fatura, itiraz listesi), `musteri-segmentasyonu-rfm` (eşitlik korumalı
  puan, 10 segmentli R×F haritası tam kapsama, iade / fatura no, kayıp sinyali, segment CSV), `cagri-merkezi-performans-raporu`
  (kısa terk, SL tanımı, FCR tekrar arama + veri sonu dışlama, Erlang C personel ihtiyacı, numara maskeleme),
  `isg-egitim-ve-muayene-takibi` (**02.04.2026 tarihli yeni İSG eğitim yönetmeliği**, RG 33212: 8/12/16 saat, 3 ay, tekrar
  3/2/1 yıl ve en az 8 saat, md. 18/19; muayene 5/3/1 yıl + özel periyot; ilkyardımcı 20/15/10 ve 3 yıl).
- **Ders:** Bash heredoc içinde kesme işareti (') sorun çıkarıyor → Write aracı. CSV'de ';' ayraçlı dosyada liste değerlerini
  '/' ile ayır. Enflasyon ortamında mutlak fiyat artışı eşiği her şeyi işaretler → medyana göre fark. İSG eğitim yönetmeliği
  2026'da yenilenmiş; mevzuat dayanaklarını her pakette güncel kaynaktan teyit et.
- **Toplam 138 paket**, kontrol 138/138.

- Devam (2026-10-10, 12. tur, 10 paket): `rakip-analizi-raporu` (birim fiyat endeksi — ambalaj boyu normalize, medyana göre
  fiyat dezavantajı, süren derin indirim, özellik eksiği / avantajı), `satis-tahmini` (naif, mevsimsel naif, HO3, Holt, Holt-Winters
  toplamsal / çarpımsal; ızgara araması; geriye dönük test WAPE ile seçim; MAD tabanlı aykırı ay + yankı kuralı),
  `teklif-hazirlama` (döviz çevrimi, aynı grupta en yüksek + gruplar arası zincirleme iskonto, müşteri dosyası ve ayrı iç kontrol:
  marj / onay), `musteri-ziyaret-plani` (vade = son ziyaret + sıklık, süpürme ile güne dağıtım, en yakın komşu + 2-opt, haversine ×
  1,3, kapasite aşımında düşük öncelik çıkar), `proforma-fatura-hazirlama` (Incoterms 2020 kuralları: deniz kuralları, navlun /
  sigorta kimde, DAT → DPU; çeki listesi koli / ağırlık / hacim; İngilizce yazıyla tutar; ayrı kontrol dosyası),
  `ihracat-evrak-kontrol-listesi` (kural tablosu: A.TR / EUR.1 / menşe şahadetnamesi, taşıma senedi, CIF-CIP sigorta, ISPM 15...;
  evraklar arası tutarlılık; ek kural / kaldır), `anket-sonuc-analizi` (frekans + hata payı, ölçek kutuları, ki-kare — gama
  fonksiyonuyla p, kritik değer tablosuyla test edildi; kelime başı kod çerçevesi; düz çizgi), `memnuniyet-nps-anketi-analizi`
  (**agent**: NPS hata payı formülü, önem × performans; tema / duygu / alıntı doğrulama; kimliğe dayalı öneriler; puan–yorum
  çelişkisi), `cagri-kalite-degerlendirmesi` (**agent**: kendi kalite formu; alıntı temsilci sözünde olmalı; anahtar ifade kod
  kontrolü; İhlal türü kriter; kritik hata → 0; doğrulanamayan → insan incelemesi), `sosyal-medya-icerik-takvimi` (**agent**:
  takvim iskeleti kodda — özel günler, anma günlerinde satış yok, kampanya dönemi; taslak denetimi: karakter sınırı, iddia,
  yasaklı ifade, doğrulanmamış oran).
- **Ders:** Özel markalı / uç değerli rakip tüm ürünleri "pahalı" gösterir → en düşüğe değil medyana göre karşılaştır. Mevsimsel
  fark tabanlı aykırı tespiti ertesi yılı da işaretler (yankı). "Yapılmaması gereken" kalite kriterlerinde alıntı mantığı ters
  çalışır (ihlal varsa alıntı). Python .lower() "İ"yi "i̇" yapar; testlerde de Türkçe katlama kullan.
- **Toplam 148 paket**, kontrol 148/148.

### KALDIĞIM YER
Sıradaki (puan 3): Akreditif Evrak Uygunluk Kontrolü, Sevkiyat Araç Planlama ve kalan2.mjs listesindeki diğerleri. Epikriz, ÖSS ve
SGK fatura ön kontrolü (puan 5) sona bırakıldı.
Sağlık/SGK faturası ve epikriz sona bırakıldı. Agent'lar gerçek API ile henüz denenmedi (kullanıcı anahtar verecek).
Listeyi üretmek için: geçici bir node betiğiyle catalog/catalog.mjs DEPTS'ten görevleri slug'layıp library/ ile karşılaştır.
İyileştirme notları: başlık eşleştirmede ı/i ve Türkçe karakter katlayan ortak bir normalizasyon (tüm paketlerde);
mizan çekirdeği para() "750.000" binlik yazımı.
Komutlar: `kontrol.py --senkronla` → `--test` → `node scripts/build.mjs` → commit + push.

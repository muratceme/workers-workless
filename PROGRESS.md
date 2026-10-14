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

- Devam (2026-10-11, 13. tur, 15 paket — tur boyu artık 15): `akreditif-evrak-uygunluk-kontrolu` (UCP 600: md. 14(c) 21 gün,
  md. 29(a) hafta sonu uzaması — ibraz için son gün de uzar, md. 18 fatura, md. 20 / 27 konşimento, md. 28 sigorta %110, md. 30
  toleranslar, md. 14(d) evraklar arası çelişki), `sevkiyat-arac-planlama` (gün + bölge grubu; palet / kg / m³; karma ve tek
  tipli denemelerden en ucuzu; küçültme; sipariş bölme; günlük araç adedi), `gise-gun-sonu-kasa-mutabakati` (kupür sayımı;
  ipuçları: tutar eşit, ters yön = 2 kat, rakam yer değişimi = 9'a bölünür, kupür katı, gişeler arası, döviz oranı),
  `musteri-evrak-kyc-eksik-kontrolu` (kural dosyası: tip × ürün × belge × geçerlilik; eksik / süresi dolmuş / yaklaşan),
  `sube-hedef-gerceklesme-panosu` (iş günü bazlı olması gereken, ay sonu tahmini, günlük gereken, tavanlı karne puanı),
  `gecikme-yaslandirma-raporu` (aralıklar, PAR30 / PAR90, müşteri bazı, geçiş matrisi), `toplu-odeme-dosyasi-kontrolu` (IBAN mod
  97 + TR kuralları, TCKN / VKN, mükerrer, kontrol toplamı, önceki dosyaya göre IBAN değişikliği), `tarihsel-riske-maruz-deger-var`
  (k = ⌊N(1−α)⌋ + 1; ES; karekök kuralı; ES katkıları toplamı = ES; geriye dönük test ve bölgeler),
  `faiz-riski-vade-uyumsuzluk-gap-analizi` (yeniden fiyatlama dilimleri; NFG etkisi = tutar × Δr × (1 − t); ED ≈ tutar × t ÷ (1 + r)
  × Δr), `gunluk-likidite-raporu` (para birimi bazında günlük / haftalık; hazır değer iskontosu; dayanma süresi; beklenen
  akışlara stres), `iade-gelen-eft-fast-analizi` (açıklamadan hata türü, Pareto, iade oranı, tekrarlayan gönderen / alıcı),
  `tfrs-9-asama-siniflandirmasi` (30 / 90 gün, yapılandırma, yakın izleme, not kademesi, TO katı; müşteri bazı; geçiş matrisi),
  `stres-testi-senaryo-hesabi` (kur / faiz / takip şokları; kâr tamponu; stres sonrası SYR; ters stres — ikiye bölme),
  `kredi-teklifi-eksik-belge-kontrolu` (kural: tip × müşteri tipi × teminat × asgari tutar × geçerlilik; şube iade metni),
  `kredi-kullandirim-on-kontrolu` (limit — sıralı düşüm, geçerlilik, ürün, vade, faiz, sözleşme, teminat tesis tutarı, şartlar).
- Ana sayfa: "İlk paketler hazır" bölümü tüm paketleri listeliyordu; 6 paketlik vitrin + "Başka paketler göster" / "Tüm
  paketlerde ara" düğmelerine çevrildi (site/assets/app.js, index.html, styles.css).
- **Ders:** Ortak yardımcılar (okuma, biçim, uyarı sayfası) her pakete aynı metinle gömülüyor; main.py'yi `#@ORTAK` işaretiyle
  yazıp bir betikle doldurmak tur başına ciddi zaman kazandırıyor (betik depo dışında, geçici klasörde). Örtük birleştirilen
  f-string'lerde `.replace(",", ".")` yalnız son parçaya uygulanır → sayıyı ayrı bir yardımcıyla biçimle. Bankacılık
  paketlerinde yasal eşikleri koda gömme: kural / eşik dosyadan veya parametreden gelsin, README "yasal ölçüm değildir" desin.
- Devam (2026-10-11, 14. tur, 15 paket, hepsi kod bloğu): `kkb-ve-risk-merkezi-raporu-ozeti` (tabloya aktarılan limit-risk satırları;
  nakdi / gayrinakdi, kullanım, banka payı, dönemsel değişim, gecikme / aşım / olumsuz kayıt bulguları; rapor şablonu okumaz),
  `sektor-karsilastirmali-rasyo-analizi` (18 rasyo; sektör değerlerini kullanıcı girer, gömülü TCMB verisi yok; sapma %, tolerans,
  çeyrek konumu, kalıp yorumlar), `teklif-talebi-bilgi-kontrolu` (kural dosyası: branş × alan × tür × seçenek × asgari × koşul;
  acenteye eksik metni), `police-uretim-kontrolu` (teklif ↔ poliçe: teminat, bedel, prim, muafiyet, tarih, özel şart; mükerrer
  üretim, üretilmemiş teklif), `sigorta-primi-hesaplama` (bedel × ‰ fiyat × gün/365 × zincirleme veya toplamsal indirim-sürprim;
  asgari prim, indirim tavanı; vergi varsayılan 0), `hasar-prim-orani-analizi` (kazanılmış prim pro-rata, ödenen + muallak;
  müşteri / acente / branş; 40-80 eşikli yenileme önerisi), `suistimal-risk-gostergeleri` (9 gösterge, puan 0-100; "gösterge
  kanıt değildir" uyarısı her çıktıda; ayar dosyası), `rucu-potansiyeli-tespiti` (ödenen × karşı kusur %; muhatap; iç takip
  tarihi — zamanaşımı hesabı DEĞİL), `tarife-etki-simulasyonu` (kural = alan × değer × değişim × kayıp; azami artış; hedef H/P
  için gerekli değişim; statik as-if), `kumul-risk-kontrolu` (bölge kümülü × PML % ÷ kapasite; lokasyon kümülü; yürürlük
  filtresi), `acente-performans-karnesi` (5 gösterge doğrusal puan, ağırlıklı karne, A-D not; eksik göstergenin ağırlığı
  dağıtılır), `prim-alacak-yaslandirma` (taksit bazında kalan alacak, aralıklar, takip listesi), `police-yenileme-takibi`
  (kapsam −15 / +30 gün; müşteri başına tek mesaj taslağı, şablon; iletişim izni işareti), `musteri-hasar-dosyasi-takibi`
  (dosya yaşı, hareketsiz gün, öncelik, aşamaya göre aksiyon, evrak hatırlatma metni), `vardiya-uretim-raporu` (sağlam adet,
  gerçekleşme, fire, kullanılabilirlik, duruş nedenleri, vardiya raporu metni; OEE değil).
  Dersler: Git Bash `sed` ile Türkçe karakterli satır değiştirme güvenilmez → Edit aracı veya Python kullan; `main()` içinde
  argparse değişkeni `a` iken `for a in ...` döngüsü onu ezer (acente karnesinde yakalandı) → döngü değişkenine başka ad ver;
  kur.py artık çıktıyı UTF-8 yazıyor (cp1254 hatası giderildi). Bu turda internetten kaynak teyidi yapılmadı: TTK 1427 (45 gün),
  halefiyet ve BSMV atıfları README'lerde genel ifadeyle ve "hukuk birimine danışın" notuyla geçiyor.
- Devam (2026-10-11 / 12, 15. tur, 15 paket, hepsi kod bloğu): `fire-ve-verim-analizi` (standart üstü fire kayıt bazında
  toplanır), `kapasite-darbogaz-analizi` (yükleme %, sezgisel kaydırma önerisi, termin riski), `olcum-verisi-uygunluk-kontrolu`
  (nokta = ondalık, binlik yok; sınıra yakın), `mtbf-ve-mttr-analizi` (çalışma süresi ÷ arıza sayısı), `zaman-etudu-ve-standart-sure`
  (performans, tolerans, frekans, gerekli gözlem sayısı), `hat-dengeleme` (RPW / LCT), `test-sonuclari-spesifikasyon-karsilastirmasi`
  (prototipler arası marj eğilimi), `numune-olcu-kontrolu` (geniş ölçü tablosu, kritik ölçü), `kesim-foyu-ve-fire-raporu`
  (beden-renk kesim %, plana göre kesim firesi), `fason-hakedis-hesaplama` (tolerans adedi, malzeme kesintisi; KDV / tevkifat
  varsayılan 0, kullanıcı verir), `final-kontrol-raporu` (hata sayısı ↔ kullanıcının girdiği kabul sayıları; plan üretmez),
  `denetim-bulgu-ve-duzeltici-faaliyet-takibi` (CAP takip durumu; standardı yorumlamaz), `karbon-ayak-izi-hesaplama` (emisyon
  faktörü kullanıcıdan; örnek faktörler "doğrulanmamış" etiketli), `musteri-haftalik-durum-raporu` (TR / EN rapor metni),
  `kategori-performans-raporu` (marj, stok gün, GMROI, ABC, işaretler).
  Dersler: liste üretecinin koşulunda `dict.pop()` her adımda çağrılır → önce değişkene al; "sevk" gibi kısa anahtar sözcük
  "Sevke hazır"ı da yakalar → anahtarı uzat; elle hesaplanan sezgisel sonuçları (kaydırma, hat dengeleme) koşturup doğrula.
  Bu turda da internetten kaynak teyidi yapılmadı.
- Devam (2026-10-12, 16. tur, 15 paket, hepsi kod bloğu): `kampanya-etki-analizi` (beklenen adet düz ortalama; kanibalizasyon
  aynı kategoriyle sınırlı; başabaş adet), `magazalar-arasi-transfer-onerisi` (kapsama haftası; sezgisel eşleştirme, aynı bölge
  önce), `sepet-birliktelik-analizi` (destek, güven, lift; yalnız ikili çiftler), `envanter-kayip-risk-skoru` (puan = oran ÷ eşik
  × 50; eşik ve ağırlıklar varsayılan), `kanallar-arasi-stok-ve-fiyat-kontrolu` (kanal fiyat farkı ayarı; kanalda olmayan ürün),
  `e-ticaret-satis-performans-raporu` (pazartesi-pazar hafta; önceki haftayla kıyas), `geciken-siparis-takibi` (termin dosyadan
  veya kanal süresinden; iş günü seçeneği; gün düzeyinde), `pazaryeri-reklam-sponsorlu-urun-analizi` (ACOS ↔ marj; "Veri az"
  sınıfı), `randevu-doluluk-ve-gelmeme-analizi` (fazla randevu = kapasite × p ÷ (1 − p) × emniyet, üst sınırlı; üç koşul),
  `taburcu-dosyasi-eksik-kontrolu` (kural dosyası: belge + koşul sütunu + diğer adlar), `kesinti-ve-itiraz-takibi` (kelime
  eşleşmesiyle neden kategorisi; itiraz süresi kullanıcıdan), `medula-ret-ve-kesinti-analizi` (Pareto, tekrarlayan hata ve
  hata-işlem çifti; SUT kuralı değerlendirmez), `kalite-gostergeleri-hesaplama` (formüller tanım dosyasından; örnek tanımlar
  "ÖRNEK TANIM" etiketli), `tedavi-fiyat-teklifi-hazirlama` (TR / EN teklif metni; tek para birimi; KDV hesabı yok),
  `magaza-vardiya-planlama` (sezgisel atama; 45 / 11 / 11 saat ve 6 gün varsayılan, parametreli).
  Dersler: Bash aracında Türkçe karakterli çok dosyalı heredoc "unexpected EOF" verdi → dosyaları Write aracıyla yaz, üretilecek
  örnek veriyi küçük Python betiğiyle üret; gün adı kısaltmasında "Pazar", "Pazartesi"nin ön eki → önce tam eşleşme, sonra tekil
  ön ek; iki sözlüğü `|` ile birleştirirken aynı anahtar (ad ↔ sayı) ezilir → alan adlarını ayır.
  Bu turda da internetten kaynak teyidi yapılmadı: vardiya paketindeki İş Kanunu sınırları, sağlık paketlerindeki SKS / Medula /
  sağlık turizmi atıfları genel ifadeyle ve "doğrulayın" notuyla geçiyor; kodlar hiçbir mevzuat değerini sabit varsaymıyor.
- Devam (2026-10-12, 17. tur, 15 paket, hepsi kod bloğu): `yuk-planlama-ldm-ve-palet` (zemin yeri × boy × en ÷ 2,4; araç tipleri
  dosyadan), `sefer-karlilik-hesabi` (sefere yazılan masraflar; km başına), `parsiyel-hakedis-ve-karlilik-hesabi` (araç maliyeti
  LDM / kg / gelire göre dağıtılır; kuruş farkı son yükte), `navlun-teklif-hesaplama` (hava 1:6000 + 0,5 kg yuvarlama; deniz W/M;
  tarife dosyadan), `konteyner-demuraj-ve-detention-takibi` (gün dilimli tarife; hat + tip önceliği; başlangıç günü dahil /
  hariç), `ithalat-vergi-hesaplama` (GV, İGV, KKDF, ÖTV, KDV matrah sırası; oranlar kullanıcıdan), `depolama-hizmeti-
  faturalandirma` (gün sonu stokuyla palet-gün; asgari fatura), `yakit-tuketim-analizi` (ilk alım hariç litre ÷ km; şüpheli
  dilim), `surucu-calisma-ve-dinlenme-suresi-kontrolu` (günlük özetten 9 / 10, 4,5, 11 / 9, 56, 90 saat kontrolleri),
  `metraj-cetveli-kontrolu` (çarpım yeniden hesaplanır; virgül ve nokta ondalık), `birim-fiyat-analizi` (rayiç eşleştirme; %25
  kâr ve genel gider parametre), `proje-maliyet-kontrol-raporu` (kalem bazında EV, EAC; iki yöntem), `fiyat-farki-hesaplama`
  (Pn = Σ katsayı × endeks oranı; F = An × B × (Pn − 1)), `is-programi-kritik-yol-analizi` (FS + gecikme; 5 / 6 / 7 günlük hafta),
  `kazanilmis-deger-analizi` (SPI, CPI, üç EAC, TCPI).
  Dersler: ölçü, miktar, katsayı ve endeks gibi alanlarda `para()` "12.500"ü 12500 okur → bu alanlarda virgül ve noktayı ondalık
  sayan yerel çözücü kullan (metraj, birim fiyat, fiyat farkı); yuvarlanmış orana göre eşik testi yazarken sınırın hemen
  ötesindeki değer yuvarlamayla sınıra düşebilir (900,01 → %10,00) → testte belirgin değer kullan; örnek dosya başlığına
  parantezli birim eklenince ("Hedef Tüketim (L/100 km)") takma ad listesine katlanmış halini de ekle.
  Bu turda da internetten kaynak teyidi yapılmadı. En çok doğrulama isteyenler: ithalat vergisinde matrah sırası, fiyat farkı
  formülü ve B = 0,90, AETR sınırları, hava kargoda 0,5 kg yuvarlama. Hepsi parametreli ve README'lerde "doğrulayın" notlu.
- **Toplam 223 paket**, kontrol 223/223.

### Tur 18 (15 kod bloğu)
provizyon-ve-sigorta-bilgisi-on-kontrolu, is-ilerleme-raporu, teklif-birim-fiyat-karsilastirmasi, geri-cagirma-izlenebilirlik-simulasyonu,
analiz-sonucu-spesifikasyon-kontrolu, parti-verim-ve-fire-raporu, recete-maliyet-hesaplama, bayi-ve-distributor-satis-raporu,
acente-kontenjan-allotment-takibi, rezervasyon-pick-up-raporu, dinamik-fiyat-onerisi, oda-temizlik-atama-plani,
menu-muhendisligi-analizi, porsiyon-maliyeti-hesaplama, mukellef-evrak-eksik-takibi.
Toplam 238 paket (206 kod + 32 agent), kontrol 238/238.
Dersler: miktar/ölçüm için yerel `olcu()` (virgül ve tek nokta ondalık, "1.250,5" binlik) kullanıldı, para için `para()`;
bölge/ay/toplam düzeyinde karşılaştırmalar (büyüme, geçen yıl, bütçe) yalnız karşılaştırma verisi olan satırlar üzerinden yapılmalı;
test yardımcı lambda'larında veri ve ayar kwargs'ları ayrılmalı; ROUND_FLOOR/ROUND_CEILING ortak importta yok, ayrıca import edilir.
İnternet teyidi yapılmadı: TCKN algoritması, Kasavana-Smith %70 eşiği, release sayımı, varsayılan parametreler (hedef doluluk 85,
maliyet oranı 30, temizlik 30/20 dk) README'lerde varsayım olarak işaretli.

### Tur 19 (15 kod bloğu)
mukellef-beyanname-takvimi, gecici-vergi-hesaplama, durusma-ve-kesin-sure-takvimi, icra-dosyasi-faiz-ve-alacak-hesabi,
arac-satis-teklifi-hesaplama, ikinci-el-arac-fiyat-degerleme, garanti-basvuru-dosyasi-kontrolu, yedek-parca-siparis-onerisi,
kazanim-bazli-basari-raporu, ise-alim-surec-raporu, sgk-giris-cikis-bildirge-kontrolu, yillik-egitim-plani,
egitim-degerlendirme-raporu, tahsilat-ve-odeme-listesi, amortisman-hesaplama.
Toplam 253 paket (221 kod + 32 agent), kontrol 253/253.
Notlar: durusma-ve-kesin-sure-takvimi, dava-ve-durusma-takvimi paketinin aynı çekirdekle farklı rol için kopyasıdır (çekirdek
değişirse ikisi birlikte güncellenmeli). Yasal oran ve günler koda gömülmedi: beyanname günleri, geçici vergi oranı, faiz oranları,
KKDF/BSMV kullanıcıdan geliyor. Koddaki yasal varsayımlar (README'de işaretli, teyit edilmedi): SGK giriş bildirgesi 1 gün önce,
çıkış 10 gün; azalan bakiyeler oranı normalin 2 katı ve en çok %50; kıst amortisman; ödeme mahsup sırası masraf → faiz → asıl.
Dersler: nested f-string'de aynı tırnak Python 3.10/3.11'de hata verir (ayrı değişkene al); argparse negatif sayıyı
"--ad=-5" biçiminde ister; test yardımcılarında .lower() yerine katla() (İ harfi).

### Tur 20 (15 kod bloğu)
ay-sonu-kapanis-kontrol-listesi, denetim-hazirlik-dosyasi, banka-limit-ve-risk-raporu, odeme-onceliklendirme,
kur-riski-pozisyon-raporu, donanim-zimmet-envanteri, yedekleme-dogrulama-raporu, kapasite-raporu, bt-proje-portfoy-durumu,
satin-alma-fiyat-gecmisi-analizi, tedarikci-risk-skoru, satin-alma-tasarruf-raporu, dava-risk-ve-karsilik-raporu,
sozlesme-envanteri-ve-sure-takibi, arac-filosu-takibi.
Toplam 268 paket (236 kod + 32 agent), kontrol 268/268.
Notlar: bu grup genel departman görevleri; yasal oran yok, eşiklerin hepsi parametre ve README'de "yönetim tercihi" notlu
(kullanım %90, RAG 0,95 / 0,85, risk skoru ağırlıkları 30-25-25-20, kapasite %90 / 30 / 90 gün, yedek toleransı 6 saat).
Teyit edilmemiş içerik varsayımları: dava raporunda TMS 37 sınıflaması (muhtemel > %50; "uzak" için %10 çalışma varsayımı),
araç filosunda kış lastiği dönemi (koda gömülmedi; örnek çalıştırmada 01.12 – 01.04 gösterim amaçlı, güncel düzenleme farklı
olabilir), ödeme önceliklendirmede sıra kuralı (zorunlu → önem → gecikmiş → gecikme maliyeti) tasarım tercihi.
Dersler: özet/kırılım adlarını sıralamadan ÖNCE üret (ilk görülen yazım girdi sırasına göre kalsın); yuvarlanmış orana göre
değil ham değere göre eşik karşılaştır (kapasite: son ≥ eşik değeri); sıfır ilerlemede endeks (MPE) hesaplama, ayrı uyarı ver;
"aşım" koşulunda tasarruf < 0 şartını da ara (eşik 0 iken sıfır tasarruf aşım sayılmasın); tarih-saat için yerel `zaman()`.

### Tur 21 (15 kod bloğu)
demirbas-ve-zimmet-takibi, pazarlama-butcesi-takibi, teklif-takip-listesi, musteri-risk-ve-limit-takibi,
fiyat-listesi-guncelleme, satis-primi-hesaplama, periyodik-kontrol-takibi, is-kazasi-ve-ramak-kala-analizi,
dahilde-isleme-izin-belgesi-takibi, stok-yaslandirma-raporu, depo-yerlesim-abc-onerisi, nakliye-maliyet-analizi,
teslimat-performansi-otif-raporu, kiymetli-evrak-sayim-kontrolu, vadesi-dolan-mevduat-takibi.
Toplam 283 paket (251 kod + 32 agent), kontrol 283/283.
Notlar: internetten kaynak teyidi yapılmadı. Eşikler parametre ve README'de "yönetim tercihi / çalışma varsayımı" notlu
(teklif takip 7 gün, limit kullanımı %90 ve gecikme 30 gün, stok 90/180 gün, ABC 80/95, nakliye medyan +%30, OTIF hedef %95,
mevduat penceresi 15 gün / acil 3 gün / 1.000.000 TL).
Teyit edilmemiş içerik varsayımları (uzman bakmalı): DİİB kapatma başvurusu süresi "süre sonu + 3 ay" (--kapatma-ay ile
değişir; 2006/12 sayılı Tebliğ'den hatırlanan süre, güncel metinle doğrulanmadı); iş kazası oranlarında çarpanlar (sıklık
1.000.000, ağırlık 1.000; SGK istatistikleri farklı çarpan kullanabilir, parametre); periyodik kontrol örnek verisindeki
süreler (kaldırma 3 ay, basınçlı kap / tesisat / yangın 12 ay, söndürücü 6 ay) yalnız gösterim, kod süre bilmez;
satış primi README'sinde "prim ücrettir, SGK + gelir vergisi + damga vergisi kesilir" ifadesi (oran verilmedi).
Tasarım tercihleri: müşteri risk kararı (Sevk etme / Onay gerekli / Uygun) müşteri düzeyinde; fiyat listesinde kural sırası
yuvarla → indirim yapma → tavan → asgari fiyat → asgari değişim; satış priminde tahsilat esası = prim × min(tahsilat/satış, 1);
OTIF sipariş bazında ve "bütün satırlar" kuralıyla; DİİB'de süre sonundan sonraki hareket sayılmaz.
Dersler: Bash heredoc ile çok dosyalı paket yazımı tırnak yüzünden kırılıyor → dosyaları Write ile yaz, heredoc yalnız küçük
Python yamaları için; argparse help metninde "%" → "%%"; sirala() kod noktasına göre sıralar ("İ" ve "Ç" Latin harflerden
sonra gelir, test beklentisini buna göre yaz); "(yok)" etiketli grup katla ile sıralanınca araya girer → açıkça sona al;
bölme sonrası yukarı yuvarlamadan önce 6 haneye yuvarla (46,4999… → 46,50 hatası olmasın); marj koruma yolunda
fiyat × yeni maliyet ÷ eski maliyet kullan (iki kez bölme yapma).

### Tur 22 (15 kod bloğu)
cek-takas-kontrol-listesi, sube-operasyonel-kontrol-listesi, teminat-takip-listesi, sube-risk-ve-gecikme-listesi,
kredi-portfoy-yogunlasma-analizi, tahsis-surec-performans-raporu, yapilandirma-takip-listesi,
takibe-aktarilacak-krediler-listesi, tahsilat-ve-avukat-performans-raporu, doviz-pozisyonu-raporu,
mevduat-maliyeti-ve-kredi-getirisi-analizi, odeme-plani-hesaplama, uyum-programi-gosterge-raporu,
kontrol-testi-orneklem-secimi, kontrol-bulgu-takibi.
Toplam 298 paket (266 kod + 32 agent), kontrol 298/298.
Notlar: bu tur bankacılık ağırlıklı (şube, tahsis, izleme-takip, hazine, operasyon, uyum, iç kontrol). İnternetten kaynak teyidi
yapılmadı. Yasal oran koda gömülmedi; hepsi parametre ve README'de "çalışma varsayımı / politikanız" notlu.
Teyit edilmemiş içerik varsayımları (uzman bakmalı):
- Çek takas: ibraz süresi varsayılanları 10 gün (aynı yer) / 1 ay (başka yer) TTK'daki genel süreler hatırlanarak kondu;
  ileri tarihli çekte "keşide tarihinden önce ibraz" kuralı dönemsel düzenleme, kod yalnız "İşleme alma" der ve doğrulatır.
- Şube risk listesi ve takibe aktarım: 90 gün takip eşiği; yapılandırma takibi: 30 gün gecikme ve 12 ay izleme süresi
  (kredilerin sınıflandırılması düzenlemesinden hatırlanan değerler; güncel metinle doğrulanmadı).
- Yoğunlaşma analizi: risk grubu / özkaynak %25 sınırı ve %10 büyük kredi eşiği (Bankacılık Kanunu'ndan hatırlanan genel
  oranlar); yasal hesaptaki dikkate alınma oranları ve istisnalar uygulanmıyor, README "yasal raporlama için kullanmayın" diyor.
- Döviz pozisyonu: limit yüzdesi varsayılmadı (kullanıcı verir); yasal YPNGP hesabının özel kuralları yok.
- Ödeme planı: BSMV ve KKDF varsayılanı 0, kullanıcı girer; örnek dosyada %15/%15 ve %5/%0 yalnız gösterim. Vergili oran
  rv = r × (1 + BSMV + KKDF) ile anüite; eşit aylık dönem varsayımı (gün sayımı yok); yıllık maliyet oranı hesaplanmıyor.
- Teminat takibi: dikkate alma oranları ve ekspertiz yenileme 12 ay tamamen kullanıcı politikası; örnek oranlar kurgusal.
Tasarım tercihleri: takibe aktarımda evrak listesi dosyanın sütunlarından okunur (bilinen sütunlar dışındaki her sütun evrak);
örneklemde yalnız rng.random() kullanılır (sürümler arası aynı sonuç), örneklem popülasyon sırasıyla yazılır; tutar yönteminde
kesin seçimden sonra aralık yeniden hesaplanır; uyum göstergelerinde yönsüz gösterge hedefle karşılaştırılmaz (ŞİB sayısına
hedef konmasın diye); döviz pozisyonu ve mevduat-kredi marjında para birimleri birbirine çevrilmeden ayrı raporlanır.
Dersler: testte elle toplam yazarken iki kez topla (bu turda üç beklenti toplama hatasıydı, kod doğruydu); "✓" gibi işaretler
katla() ile boşalır → ham metne ayrıca bak; argparse'ta değişken adı `a` ile parametre `a` çakışmasın (args kullan).

### Tur 23 (15 kod bloğu)
operasyonel-risk-kayip-veri-analizi, brans-karlilik-raporu, muallak-hasar-takibi, muallak-hasar-raporu, rucu-tahsilat-takibi,
eksper-performans-analizi, kurum-fiyat-karsilastirmasi, uretim-hedef-takibi, prim-odenmemesi-nedeniyle-iptal-listesi,
acente-portfoy-ve-komisyon-raporu, musteri-tahsilat-takibi, durus-nedeni-pareto-analizi, uretim-performans-panosu,
plan-uyum-raporu, dof-takibi.
Toplam 313 paket (281 kod + 32 agent).
Notlar: bu tur sigorta (teknik, hasar, sağlık, satış, tahsilat, acente) ve üretim (üretim, planlama, kalite) ağırlıklı.
İnternetten kaynak teyidi yapılmadı. Yasal oran ve süre koda gömülmedi; hepsi parametre.
Teyit edilmemiş içerik varsayımları (uzman bakmalı):
- Prim ödenmemesi / iptal listesi: ilk taksitte vadeden başlayan 3 aylık cayma süresi ve izleyen taksitlerde ihtarla 10 gün
  (TTK'nın prim borcunda temerrüt maddesinden hatırlanan genel kural). Madde numarası README'ye yazılmadı; zorunlu sigortalar,
  hayat ve sağlıkta farklı kurallar olabileceği notu var. Kod iptal kararı vermez, "aday" der.
- Rücu takibi: zamanaşımı süresi hesaplanmıyor, tarihi kullanıcı yazıyor; cevap ve tahsilat için 30 gün çalışma eşiği.
- Branş kârlılığı: komisyon ve gider oranının paydası kazanılmış prim (bazı raporlarda yazılan prim kullanılır, README'de
  notlu); mali tablodaki teknik kâr değil (yatırım geliri, DERK, dengeleme karşılığı yok).
- Muallak hasar raporu: "gerçekleşme oranı" (ödenen ÷ ilk muallak) yönetim göstergesi; aktüeryal yeterlilik hesabı değil.
- Operasyonel risk: Basel'in 7 olay türü / 8 iş kolu yalnız README'de anılıyor; kod sınıflandırmayı dayatmıyor.
- Üretim panosu: OEE hedefi varsayılanı %60 (politika); "dünya sınıfı %85" her süreç için ölçüt değildir notu var.
- Acente raporları: komisyon oranı varsayılmıyor; şirkete aktarılacak tutarı kullanıcı yazıyor (boşsa taksitin tamamı).
Tasarım tercihleri: oranlar her yerde toplamların oranı (ortalamaların ortalaması değil); kurum fiyat endeksi ortancaya göre
ve adet ağırlıklı; Pareto'da A sınıfı = kendisinden önceki kümülatif pay eşiğin altında olan kalem (eşiği geçiren kalem
dahil); plan uyumunda fazla üretim miktar uyumunu artırmaz ve "uygun" sayılmaz; DÖF ancak etkinlik "Etkin" ise kapalı;
üretim hedefinde yalnız yıl yazılan hedef 12 aya eşit bölünür.
Dersler: adet / miktar sütunlarında olcu() "1.000" değerini 1 okur → adet ve süre sütunlarında para() kullan (bu turda
kurum fiyat, duruş pareto ve üretim panosunda düzeltildi; önceki turlarda olcu() ile okunan adet sütunları gözden
geçirilmeli); Bash heredoc içindeki Python'da "\n" kaçışı gerçek satır sonuna dönüşüyor → test dosyası yamalarında
chr(92) ya da Edit aracı kullan.

### Tur 24 (15 kod bloğu)
kalitesizlik-maliyeti-raporu, tedarikci-kalite-karnesi, periyodik-bakim-is-emri-listesi, yedek-parca-kritik-stok-onerisi,
bakim-maliyet-ve-performans-raporu, kaizen-onerisi-degerlendirme, urun-agaci-bom-karsilastirma, lab-dip-ve-onay-takibi,
kumas-fiyat-takibi, hata-turu-pareto-analizi, gun-sonu-kasa-raporu, kasa-fark-ve-iptal-analizi, magaza-satis-performans-panosu,
fire-ve-kayip-analizi, personel-satis-performansi.
Toplam 328 paket (296 kod + 32 agent).
Notlar: bu tur kalite, bakım, Ar-Ge, tekstil ve mağaza operasyonları. İnternetten kaynak teyidi yapılmadı; yasal oran yok.
Varsayımlar (politika, kullanıcı değiştirir): tedarikçi karnesi A ≤ 1.000 ppm / B ≤ 5.000 ppm; dış hata payı %50 uyarısı;
yedek parça emniyet payı %100/%50/%25 (istatistiksel model değil, README açıkça söylüyor); planlı bakım payı hedefi %70;
kaizen eşikleri ve ağırlıkları 50/30/20; lab dip cevap süresi 7 gün; kumaş fiyat artış uyarısı %10; kasa toleransı 5 TL;
kasa analizinde "diğer kasiyerlerin" oranının 2 katı; mağaza LFL eşiği −%5, hedef %90; fire/kayıp eşiği %2; personel
priminde kod varsayılan kademe kullanmıyor (örnek veri 90:1;100:2;120:3 ile çalışıyor).
Tasarım tercihleri: kasa analizinde kasiyer kendi oranıyla değil, kendisi hariç diğerlerinin toplam oranıyla karşılaştırılır
(az kişide ortalamayı tek aykırı değer şişirmesin diye); BOM'da 0.250 ondalık, 1.000 binlik okunur (miktar_oku); bakım
planında malzeme hücresi "Ad x2, Ad" biçiminde, x'ten önce boşluk şart (10x20 gibi adlar bozulmasın); hata pareto ve kasa
analizinde kişisel veri / suistimal değildir notları README'de.
Dersler: gruplamayı sıralamadan önce yap ya da ilk görülen yazımı sakla (kaizen bölüm adı); uyarı "kim" alanı strip edilir
(test beklentisinde sondaki boşluk olmaz); yuvarlanmış ara değerden oran hesaplayıp test yazma (personel UPT endeksi).

### Tur 25 (15 kod bloğu)
bolge-magaza-karsilastirma-raporu, rakip-fiyat-karsilastirmasi, urun-gami-delist-onerisi, tedarikci-ciro-primi-hesabi,
satis-hizi-sell-through-raporu, kayip-musteri-churn-listesi, kampanya-fiyat-simulasyonu, urun-basligi-seo-kontrolu,
kargo-teslim-performansi, hasta-kabul-gun-sonu-kasa-kontrolu, servis-yatak-doluluk-raporu, kurum-bazli-gelir-ve-kesinti-raporu,
son-kullanma-tarihi-takibi, kritik-stok-seviyesi-raporu, teklif-donusum-analizi.
Toplam 343 paket (311 kod + 32 agent).
Notlar: perakende kategori, e-ticaret, hastane (kasa, yatak, kurum faturası, eczane) ve lojistik satış. İnternetten kaynak
teyidi yapılmadı. Yasal dayanak olarak yalnız churn README'sinde 6563 sayılı Kanun / İYS onay şartı anıldı; KDV oranı
varsayılanı %20 (yalnız kampanya simülasyonunda, KDV sütunu boşsa; --kdv ile değişir). Yatak doluluk / ortalama kalış
formülleri yaygın hastane istatistiği tanımı; README kurum ve SB tanımıyla doğrulanmasını söylüyor.
Varsayımlar (politika): bölge sapma %20, LFL −%5; rakip endeks bandı ±%5, tarama geçerliliği 30 gün, yarı/iki kat şüpheli
eşleşme; delist kuyruğu %10, raf verimi 50, yeni ürün 90 gün; sell-through %40, 4 hafta, stok günü 7/60; churn 180 gün kayıp,
ortalama aralığın 2 katı risk, tek alışveriş 90 gün; kampanya min. marj %10, komisyon matrahı KDV dahil; başlık 30-100
karakter + varsayılan yasaklı ifade listesi; kargo hedef 3 gün, %90 zamanında, %1 hasar, %0,5 kayıp; kasa toleransı 1 TL;
doluluk %90/%50; kesinti %5, tahsil 90 gün; SKT yakın 90 gün (FEFO); kritik stok emniyet 7 gün, kapsam 30 gün;
teklif kazanma eşiği %20 (en az 5 sonuçlanmış teklif), bekleme 14 gün.
Tasarım tercihleri: rakip fiyatında şüpheli eşleşme ve kampanyalı tarama endekse girmez; başlık önerisi sıra doğruysa
özgün kelime sırasını korur, özellik eşleşmesinde ek almış hâl sayılır (pamuk → pamuklu); kritik stokta sipariş önerisi yalnız
stoksuz / kritik altındakilere (sipariş noktası mantığı); SKT'de ay-yıl yazımı ayın son günü.
Dersler: Bash heredoc'u bazı içeriklerde (çok satırlı CSV + test) "unexpected EOF" verdi; uzun dosyaları Write aracıyla yaz.
Python 3.10 hedefinde f-string içinde aynı tırnakla iç içe f-string kullanma (3.12 özelliği).

### KALDIĞIM YER
Sıradaki (puan 2, kod): Hasta Seyahat ve Tedavi Takvimi, Hastane Operasyon Panosu, Hekim Performans Raporu, Güvenlik
Raporlama Sistemi Olay Analizi, SKS Öz Değerlendirme Takibi, Beton Numune ve Kırım Takibi, Şantiye Puantajı, Malzeme Talep ve
Sevk Takibi, İmalat Kontrol Formu Takibi, Denetim Hazırlık Kontrol Listesi, Besin Değeri Hesaplama, Raf Ömrü Takibi, Temizlik ve
Sanitasyon Kayıt Kontrolü, Duyusal Test Sonuç Analizi, Zincir Market Listeleme Takibi, otel (giriş-çıkış, no-show, grup teklifi,
kanal maliyeti, oda arıza, F&B maliyet), e-Defter berat takibi… (tam liste için kalan2.mjs; proje klasöründen çalıştır).
Agent adayları (puan 2): Teminat Mektubu Metin Kontrolü, Müşteri İhtiyaç Analizi, Vardiya Devir Notu Özeti, Mağaza Ziyaret
Raporu, Vitrin Yönergesi, Renk Paleti, Trend Raporu Özeti, Randevu Hatırlatma Mesajları, Tıbbi Rapor Yazım Kontrolü.
Agent olarak bekleyenler: Şüpheli İşlem Bildirim Taslağı, Hasar İhbarı Ön Kayıt, Müşteri Şikâyeti Kök Neden,
Arıza Kaydı Sınıflandırma, Ar-Ge Proje Raporu Taslağı, Fit Comments Özeti, İade Nedeni Analizi, Ürün Değerlendirme Analizi,
Planogram Uyum, Hasta Şikâyet Sınıflandırma, Hasta Memnuniyet Anketi Analizi, Müşteri Yük Durum Bilgilendirmesi, Konşimento Bilgi
Çıkarma, Günlük Şantiye Raporu, Etiket Bilgisi Kontrolü, Misafir Mesajlarına Cevap Taslağı, Yorum Cevap Taslağı, Karar ve İçtihat
Özeti. Epikriz, ÖSS ve SGK fatura (puan 5) sonda.
Kullanıcı tur boyunu 15 pakete çıkardı (15 paket → dur → /compact).
Sağlık/SGK faturası ve epikriz sona bırakıldı. Agent'lar gerçek API ile henüz denenmedi (kullanıcı anahtar verecek).
Listeyi üretmek için: geçici bir node betiğiyle catalog/catalog.mjs DEPTS'ten görevleri slug'layıp library/ ile karşılaştır.
İyileştirme notları: başlık eşleştirmede ı/i ve Türkçe karakter katlayan ortak bir normalizasyon (tüm paketlerde);
mizan çekirdeği para() "750.000" binlik yazımı.
Komutlar: `kontrol.py --senkronla` → `--test` → `node scripts/build.mjs` → commit + push.

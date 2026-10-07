/* =========================================================================
   Workers / Workless — katalog
   Sektör → Departman → Rol (çalışan) → Görev
   Görev formatı: [ad, açıklama, haftalık kazandırdığı saat]

   Görev kimliği (id) = görev adının slug'ı ve TÜM katalogda benzersizdir.
   library/code-blocks/<id>/ ve library/agents/<id>/ klasörleri bu kimlikle eşleşir.
   Bir görev birden çok sektörde görünse de kodu tek bir yerde durur.
   `shared: true` departmanlar (İK, Finans, BT) her sektörde aynıdır.
   ========================================================================= */

export const CONFIG = {
  githubUser: "muratceme",
  repo: "workers-workless",
  branch: "main",
};

export const DEPTS = {
  /* ------------------------------ ORTAK ------------------------------ */
  ik: {
    name: "İnsan Kaynakları", short: "İK", shared: true,
    desc: "İşe alımdan bordroya, insanla ilgili tekrar eden her iş.",
    roles: [
      { id: "ik-stajyeri", name: "İK Stajyeri", level: "Giriş", tasks: [
        ["CV Ön Eleme", "Gelen CV'leri ilan kriterlerine göre ilk filtreden geçirir, elenenleri gerekçesiyle listeler.", 6],
        ["Mülakat Takvimi", "Aday ve mülakatçı müsaitliklerini eşleştirip takvim davetlerini hazırlar.", 3],
        ["Aday Bilgilendirme E-postaları", "Süreç aşamasına göre kişiselleştirilmiş aday e-postaları üretir.", 2],
      ]},
      { id: "ik-uzmani", name: "İK Uzmanı", level: "Uzman", tasks: [
        ["CV Raporlama", "Yüzlerce CV'yi okuyup deneyim, eğitim ve yetkinliklere göre tek bir Excel raporunda toplar.", 8],
        ["CV – İlan Eşleştirme Skoru", "Her aday için ilana uygunluk skoru ve kısa gerekçe üretir.", 5],
        ["İşe Giriş Evrak Kontrolü", "Yeni çalışan evraklarının eksiksiz ve geçerli olup olmadığını kontrol eder.", 3],
        ["Bordro Veri Doğrulama", "Puantaj ve bordro verisindeki tutarsızlıkları ay kapanmadan yakalar.", 4],
      ]},
      { id: "ik-muduru", name: "İK Müdürü", level: "Yönetici", tasks: [
        ["Personel Devir Analizi", "Ayrılan çalışan verisinden departman bazlı devir oranı ve risk sinyalleri çıkarır.", 3],
        ["Performans Değerlendirme Özeti", "Yönetici değerlendirmelerini çalışan başına tek sayfalık özete çevirir.", 5],
        ["Ücret Bandı Karşılaştırması", "İç ücret verisini piyasa bantlarıyla kıyaslar, sapmaları işaretler.", 2],
      ]},
    ],
  },
  finans: {
    name: "Finans & Muhasebe", short: "FİN", shared: true,
    desc: "Fatura, mutabakat, bütçe ve raporlama döngüsü.",
    roles: [
      { id: "muhasebe-elemani", name: "Muhasebe Elemanı", level: "Giriş", tasks: [
        ["Fatura Okuma & Kayıt", "PDF / e-Fatura dosyalarını okuyup muhasebe kayıt şablonuna dönüştürür.", 10],
        ["Banka Mutabakatı", "Banka ekstresi ile defter kayıtlarını eşleştirir, açık kalemleri listeler.", 6],
        ["Masraf Kategorileme", "Çalışan masraf fişlerini hesap planına göre sınıflandırır.", 3],
      ]},
      { id: "finansal-analist", name: "Finansal Analist", level: "Uzman", tasks: [
        ["Nakit Akışı Tahmini", "Geçmiş tahsilat/ödeme verisinden 13 haftalık nakit tahmini üretir.", 4],
        ["Bütçe Sapma Raporu", "Bütçe ile gerçekleşeni karşılaştırıp sapma nedenlerini özetler.", 5],
        ["Yönetim KPI Raporu", "Aylık yönetim raporunun tablo ve grafiklerini otomatik hazırlar.", 6],
      ]},
      { id: "finans-muduru", name: "Finans Müdürü", level: "Yönetici", tasks: [
        ["Konsolide Finansal Rapor", "Şirketlerin mizanlarını birleştirip konsolide tabloları üretir.", 6],
        ["Senaryo & Duyarlılık Analizi", "Kur, faiz ve satış varsayımlarına göre kârlılık senaryoları çıkarır.", 3],
        ["Vergi Takvimi Takibi", "Beyanname ve ödeme tarihlerini takip edip hatırlatma listesi oluşturur.", 1],
      ]},
    ],
  },
  bt: {
    name: "Bilgi Teknolojileri", short: "BT", shared: true,
    desc: "Destek talepleri, sistemler, lisanslar ve güvenlik.",
    roles: [
      { id: "helpdesk-uzmani", name: "Helpdesk Uzmanı", level: "Giriş", tasks: [
        ["Ticket Sınıflandırma", "Gelen destek taleplerini kategori ve önceliğe göre otomatik etiketler.", 5],
        ["SSS Yanıtlayıcı", "Sık sorulan sorulara bilgi bankasından cevap taslağı üretir.", 6],
        ["Hesap Açma / Kapama", "İşe giriş-çıkış listesine göre kullanıcı hesap işlemlerini hazırlar.", 3],
      ]},
      { id: "sistem-yoneticisi", name: "Sistem Yöneticisi", level: "Uzman", tasks: [
        ["Log Anomali Tespiti", "Sunucu loglarında olağandışı örüntüleri bulur ve özetler.", 4],
        ["Yedekleme Doğrulama", "Yedekleme işlerinin sonuçlarını kontrol edip başarısızları raporlar.", 2],
        ["Lisans Envanteri", "Kurulu yazılımları lisans kayıtlarıyla karşılaştırır.", 2],
      ]},
      { id: "bt-muduru", name: "BT Müdürü", level: "Yönetici", tasks: [
        ["BT Maliyet Raporu", "Bulut, lisans ve donanım harcamalarını birim bazında raporlar.", 3],
        ["Siber Risk Envanteri", "Varlık listesi ve zafiyet taramasından risk matrisi oluşturur.", 2],
      ]},
    ],
  },

  /* ---------------------------- BANKACILIK ---------------------------- */
  krediler: {
    name: "Krediler", short: "KRD",
    desc: "Kredi analizi, tahsis ve portföy riski.",
    roles: [
      { id: "kredi-analisti", name: "Kredi Analisti", level: "Uzman", tasks: [
        ["Mali Tablo Analizi", "Bilanço ve gelir tablosu PDF'lerinden finansal oranları çıkarır, yorumlar.", 8],
        ["Kredi Skor Kartı", "Firma verisinden kural tabanlı ön kredi skoru hesaplar.", 4],
        ["Teminat Değerleme Özeti", "Ekspertiz raporlarından teminat değer ve risk özeti çıkarır.", 3],
        ["İstihbarat Raporu Derleme", "Farklı kaynaklardaki firma istihbaratını tek rapora toplar.", 4],
      ]},
      { id: "kredi-tahsis-muduru", name: "Kredi Tahsis Müdürü", level: "Yönetici", tasks: [
        ["Komite Sunum Özeti", "Kredi dosyasını komite için tek sayfalık karar özetine dönüştürür.", 4],
        ["Portföy Yoğunlaşma Analizi", "Sektör ve grup bazında risk yoğunlaşmasını görselleştirir.", 2],
        ["Erken Uyarı Sinyalleri", "Ödeme davranışı ve hesap hareketlerinden erken uyarı listesi üretir.", 3],
      ]},
    ],
  },
  uyum: {
    name: "Uyum & Mevzuat", short: "UYM",
    desc: "MASAK, KYC, yaptırım taramaları ve denetim.",
    roles: [
      { id: "uyum-uzmani", name: "Uyum Uzmanı", level: "Uzman", tasks: [
        ["Şüpheli İşlem Tespiti", "İşlem verisinde MASAK senaryolarına uyan örüntüleri işaretler.", 6],
        ["KYC Belge Doğrulama", "Müşteri kimlik ve adres belgelerinin tutarlılığını kontrol eder.", 5],
        ["Yaptırım Listesi Taraması", "Müşteri listesini güncel yaptırım listeleriyle bulanık eşleştirir.", 3],
      ]},
      { id: "uyum-muduru", name: "Uyum Müdürü", level: "Yönetici", tasks: [
        ["Mevzuat Değişiklik Takibi", "Resmî Gazete ve BDDK duyurularını tarayıp ilgili değişiklikleri özetler.", 3],
        ["Denetim Bulgu Takibi", "Açık denetim bulgularını sorumlu ve tarihe göre takip eder.", 2],
      ]},
    ],
  },
  sube: {
    name: "Şube Operasyonları", short: "ŞUB",
    desc: "Gişe, müşteri temsilciliği ve şube yönetimi.",
    roles: [
      { id: "gise-yetkilisi", name: "Gişe Yetkilisi", level: "Giriş", tasks: [
        ["Dekont Sınıflandırma", "Taranmış dekont ve talimatları türüne göre ayırıp arşivler.", 3],
        ["Müşteri Talep Yönlendirme", "Gelen talepleri doğru birime yönlendirme önerisi üretir.", 2],
      ]},
      { id: "musteri-temsilcisi", name: "Müşteri Temsilcisi", level: "Uzman", tasks: [
        ["Ürün Öneri Motoru", "Müşteri profiline göre uygun bankacılık ürünlerini önerir.", 3],
        ["Görüşme Notu Özeti", "Müşteri görüşme notlarını CRM formatında özetler.", 3],
        ["Kampanya Hedef Listesi", "Kampanya kriterlerine uyan müşteri listesini çıkarır.", 2],
      ]},
      { id: "sube-muduru", name: "Şube Müdürü", level: "Yönetici", tasks: [
        ["Şube Performans Panosu", "Günlük hedef / gerçekleşme verisini tek panoda toplar.", 3],
        ["Hedef Gerçekleşme Raporu", "Temsilci bazında hedef gerçekleşme ve sıralama raporu üretir.", 2],
      ]},
    ],
  },

  /* ------------------------------ SİGORTA ----------------------------- */
  hasar: {
    name: "Hasar", short: "HSR",
    desc: "Hasar dosyası, ekspertiz ve suistimal kontrolü.",
    roles: [
      { id: "hasar-asistani", name: "Hasar Asistanı", level: "Giriş", tasks: [
        ["Hasar Fotoğrafı Ön Değerlendirme", "Araç/konut hasar fotoğraflarından hasarlı bölge ve şiddet tahmini çıkarır.", 5],
        ["Evrak Eksiklik Kontrolü", "Hasar dosyasındaki eksik evrakları tespit edip talep listesi oluşturur.", 4],
      ]},
      { id: "hasar-uzmani", name: "Hasar Uzmanı", level: "Uzman", tasks: [
        ["Hasar Dosyası Özeti", "Onlarca sayfalık dosyayı karar için tek sayfaya indirir.", 6],
        ["Suistimal Risk Skoru", "Dosyadaki şüpheli örüntülere göre suistimal risk skoru üretir.", 4],
        ["Rücu Potansiyeli Tespiti", "Kusur ve sorumluluk bilgisinden rücu edilebilir dosyaları işaretler.", 2],
      ]},
      { id: "hasar-muduru", name: "Hasar Müdürü", level: "Yönetici", tasks: [
        ["Muallak Hasar Raporu", "Açık dosyaların yaş ve tutar dağılımını raporlar.", 3],
        ["Eksper Performans Analizi", "Eksper bazında süre, maliyet ve itiraz oranlarını karşılaştırır.", 2],
      ]},
    ],
  },
  akturya: {
    name: "Aktüerya", short: "AKT",
    desc: "Fiyatlama, rezerv ve risk modelleri.",
    roles: [
      { id: "akturer", name: "Aktüer", level: "Uzman", tasks: [
        ["Prim Hesaplama Motoru", "Risk faktörlerinden tarife primini hesaplayan parametrik motor.", 5],
        ["Frekans – Şiddet Analizi", "Hasar sıklığı ve ortalama hasar tutarını segment bazında çıkarır.", 4],
        ["Rezerv Tahmini (Chain Ladder)", "Gelişim üçgeninden IBNR rezerv tahmini üretir.", 4],
      ]},
      { id: "kidemli-akturer", name: "Kıdemli Aktüer", level: "Yönetici", tasks: [
        ["Tarife Senaryo Simülasyonu", "Fiyat değişikliğinin portföy ve kârlılığa etkisini simüle eder.", 3],
        ["Model Doğrulama Raporu", "Fiyatlama modellerinin geri test ve stabilite raporunu hazırlar.", 3],
      ]},
    ],
  },
  police: {
    name: "Satış & Poliçe", short: "PLÇ",
    desc: "Teklif, yenileme ve acente yönetimi.",
    roles: [
      { id: "satis-temsilcisi", name: "Satış Temsilcisi", level: "Giriş", tasks: [
        ["Poliçe Teklif Karşılaştırma", "Farklı teminat seçeneklerini müşteri için anlaşılır tabloya döker.", 4],
        ["Yenileme Hatırlatma", "Vadesi yaklaşan poliçeler için kişisel hatırlatma mesajları hazırlar.", 3],
        ["İhtiyaç Analizi Formu", "Görüşme notlarından müşteri ihtiyaç analizi çıkarır.", 2],
      ]},
      { id: "acente-yoneticisi", name: "Acente Yöneticisi", level: "Yönetici", tasks: [
        ["Acente Performans Raporu", "Acente bazında üretim, hasar/prim ve tahsilat performansı.", 3],
        ["Üretim Hedef Takibi", "Branş bazında hedef gerçekleşmeyi haftalık takip eder.", 2],
      ]},
    ],
  },

  /* ------------------------------ ÜRETİM ------------------------------ */
  planlama: {
    name: "Üretim Planlama", short: "PLN",
    desc: "Çizelgeleme, malzeme ihtiyacı ve stok.",
    roles: [
      { id: "uretim-planlama-uzmani", name: "Üretim Planlama Uzmanı", level: "Uzman", tasks: [
        ["Haftalık Üretim Çizelgesi", "Sipariş ve kapasite verisinden makine bazlı çizelge üretir.", 8],
        ["Malzeme İhtiyaç Planı (MRP)", "Ürün ağacı ve siparişlerden malzeme ihtiyaç listesi çıkarır.", 6],
        ["Kapasite Darboğaz Analizi", "Hat bazlı yüklemeyi hesaplayıp darboğazları gösterir.", 3],
      ]},
      { id: "stok-kontrol-sorumlusu", name: "Stok Kontrol Sorumlusu", level: "Giriş", tasks: [
        ["Yeniden Sipariş Noktası", "Tüketim ve tedarik süresinden emniyet stoğu ve sipariş noktası hesaplar.", 3],
        ["Stok Yaşlandırma Raporu", "Hareketsiz ve yavaş dönen stokları listeler.", 2],
        ["Sayım Fark Analizi", "Sayım ile sistem stoğu arasındaki farkları açıklamalı raporlar.", 3],
      ]},
    ],
  },
  kalite: {
    name: "Kalite Kontrol", short: "KLT",
    desc: "Muayene, SPC, 8D ve tedarikçi kalitesi.",
    roles: [
      { id: "kalite-teknisyeni", name: "Kalite Teknisyeni", level: "Giriş", tasks: [
        ["Görsel Hata Tespiti", "Hat kamerası görüntülerinden yüzey hatalarını işaretler.", 8],
        ["Ölçüm Verisi Kaydı", "Ölçüm cihazı çıktılarını standart kayıt formatına çevirir.", 3],
      ]},
      { id: "kalite-muhendisi", name: "Kalite Mühendisi", level: "Uzman", tasks: [
        ["SPC Kontrol Grafikleri", "Proses verisinden X̄-R grafikleri ve Cpk değerleri üretir.", 4],
        ["8D Rapor Taslağı", "Uygunsuzluk bilgilerinden 8D raporunun ilk taslağını hazırlar.", 4],
        ["Şikayet Kök Neden Analizi", "Müşteri şikayetlerini kümeleyip olası kök nedenleri önerir.", 3],
      ]},
      { id: "kalite-muduru", name: "Kalite Müdürü", level: "Yönetici", tasks: [
        ["Kalite Maliyetleri Raporu", "Hurda, yeniden işleme ve iade maliyetlerini raporlar.", 2],
        ["Tedarikçi Kalite Skoru", "Giriş kalite verisinden tedarikçi skor kartı üretir.", 2],
      ]},
    ],
  },
  bakim: {
    name: "Bakım", short: "BKM",
    desc: "Arıza, periyodik ve kestirimci bakım.",
    roles: [
      { id: "bakim-teknisyeni", name: "Bakım Teknisyeni", level: "Giriş", tasks: [
        ["Arıza Kaydı Sınıflandırma", "Serbest metin arıza kayıtlarını ekipman ve arıza tipine göre etiketler.", 3],
        ["Periyodik Bakım Hatırlatıcı", "Bakım planından haftalık iş emri listesi üretir.", 2],
      ]},
      { id: "bakim-muhendisi", name: "Bakım Mühendisi", level: "Uzman", tasks: [
        ["Kestirimci Bakım Modeli", "Sensör verisinden arıza olasılığı yüksek ekipmanları tahmin eder.", 4],
        ["OEE Hesaplama", "Duruş, hız ve kalite kayıplarından OEE raporu üretir.", 3],
        ["Yedek Parça Optimizasyonu", "Arıza geçmişine göre kritik yedek parça stok seviyelerini önerir.", 2],
      ]},
    ],
  },

  /* ------------------------------ TEKSTİL ----------------------------- */
  tasarim: {
    name: "Tasarım", short: "TSR",
    desc: "Koleksiyon, desen, renk ve teknik föy.",
    roles: [
      { id: "tasarim-asistani", name: "Tasarım Asistanı", level: "Giriş", tasks: [
        ["Renk Paleti Çıkarma", "Görsellerden Pantone yakınlığıyla renk paleti çıkarır.", 2],
        ["Moodboard Derleme", "Tema anahtar kelimelerinden düzenli bir moodboard klasörü hazırlar.", 3],
      ]},
      { id: "moda-tasarimcisi", name: "Moda Tasarımcısı", level: "Uzman", tasks: [
        ["Desen Varyasyon Üretici", "Bir desenden renk ve ölçek varyasyonları üretir.", 4],
        ["Trend Raporu Özeti", "Sezon trend raporlarını uygulanabilir maddelere indirir.", 2],
        ["Teknik Föy Taslağı", "Model bilgilerinden tech pack taslağını oluşturur.", 5],
      ]},
    ],
  },
  numune: {
    name: "Numune & Kalite", short: "NMN",
    desc: "Ölçü tabloları, numune takibi ve kumaş kontrolü.",
    roles: [
      { id: "numune-sorumlusu", name: "Numune Sorumlusu", level: "Uzman", tasks: [
        ["Ölçü Tablosu Grading", "Ana bedenden beden serisine ölçü tablosu üretir.", 4],
        ["Numune Takip Çizelgesi", "Numune aşamalarını müşteri ve termine göre takip eder.", 3],
      ]},
      { id: "kumas-kalite-kontrolcu", name: "Kumaş Kalite Kontrolcü", level: "Giriş", tasks: [
        ["Kumaş Hata Tespiti", "Kumaş top görüntülerinden dokuma ve boya hatalarını işaretler.", 6],
        ["4 Puan Sistemi Raporu", "Hata kayıtlarından 4 puan sistemine göre top raporu hesaplar.", 3],
      ]},
    ],
  },
  tedarik: {
    name: "Satınalma & Tedarik", short: "TDR",
    desc: "İplik, kumaş, aksesuar ve tedarikçi yönetimi.",
    roles: [
      { id: "satinalma-uzmani", name: "Satınalma Uzmanı", level: "Uzman", tasks: [
        ["İplik & Kumaş Fiyat Takibi", "Tedarikçi fiyat listelerini toplayıp trend tablosu oluşturur.", 3],
        ["Tedarikçi Teklif Karşılaştırma", "Gelen teklifleri fiyat, termin ve kaliteye göre puanlar.", 4],
        ["Sipariş Termin Takibi", "Geciken siparişleri tespit edip tedarikçiye hatırlatma hazırlar.", 3],
      ]},
      { id: "tedarik-zinciri-muduru", name: "Tedarik Zinciri Müdürü", level: "Yönetici", tasks: [
        ["Tedarikçi Risk Skoru", "Gecikme, kalite ve finansal veriden tedarikçi risk skoru üretir.", 2],
        ["Karbon Ayak İzi Raporu", "Malzeme ve lojistik verisinden ürün bazlı karbon tahmini çıkarır.", 3],
      ]},
    ],
  },

  /* ----------------------------- PERAKENDE ---------------------------- */
  eticaret: {
    name: "E-ticaret", short: "ETC",
    desc: "Ürün içerikleri, pazaryerleri ve müşteri deneyimi.",
    roles: [
      { id: "eticaret-uzmani", name: "E-ticaret Uzmanı", level: "Uzman", tasks: [
        ["Ürün Açıklaması Üretici", "Ürün özelliklerinden SEO uyumlu açıklamalar üretir.", 8],
        ["Pazaryeri Listeleme Senkronu", "Ürün kataloğunu pazaryeri şablonlarına dönüştürür.", 5],
        ["Rakip Fiyat Takibi", "Rakip fiyatlarını izleyip fiyat farkı raporu üretir.", 3],
      ]},
      { id: "musteri-deneyimi-uzmani", name: "Müşteri Deneyimi Uzmanı", level: "Giriş", tasks: [
        ["Yorum Duygu Analizi", "Ürün yorumlarını duygu ve konuya göre sınıflandırır.", 3],
        ["İade Nedeni Sınıflandırma", "İade açıklamalarından neden kategorileri çıkarır.", 2],
        ["Hazır Cevap Önerici", "Müşteri mesajına uygun cevap taslağı önerir.", 6],
      ]},
    ],
  },
  magaza: {
    name: "Mağaza Operasyonları", short: "MĞZ",
    desc: "Raf, kasa, vardiya ve mağaza performansı.",
    roles: [
      { id: "magaza-personeli", name: "Mağaza Personeli", level: "Giriş", tasks: [
        ["Planogram Kontrolü", "Raf fotoğrafını planogramla karşılaştırıp eksikleri listeler.", 3],
        ["Gün Sonu Kasa Raporu", "POS çıktılarından gün sonu raporunu hazırlar.", 2],
      ]},
      { id: "magaza-muduru", name: "Mağaza Müdürü", level: "Yönetici", tasks: [
        ["Vardiya Planlama", "Trafik tahmini ve personel kısıtlarından vardiya planı üretir.", 4],
        ["Mağaza Satış Panosu", "Günlük satış, sepet ve dönüşüm metriklerini tek panoda toplar.", 2],
        ["Fire & Kayıp Analizi", "Fire kayıtlarından ürün ve neden bazlı kayıp raporu üretir.", 2],
      ]},
    ],
  },
  talep: {
    name: "Talep Planlama", short: "TLP",
    desc: "Satış tahmini ve dağıtım.",
    roles: [
      { id: "talep-planlama-uzmani", name: "Talep Planlama Uzmanı", level: "Uzman", tasks: [
        ["Satış Tahmini", "Geçmiş satış ve mevsimsellikten ürün-mağaza bazlı tahmin üretir.", 6],
        ["Kampanya Etki Analizi", "Kampanya dönemindeki ek satışı (uplift) ölçer.", 3],
        ["Bölgesel Dağıtım Önerisi", "Tahmin ve stoğa göre mağazalara sevk miktarı önerir.", 4],
      ]},
    ],
  },

  /* ------------------------------- SAĞLIK ------------------------------ */
  "hasta-kabul": {
    name: "Hasta Kabul", short: "HKB",
    desc: "Randevu, kayıt ve hasta iletişimi.",
    roles: [
      { id: "hasta-kabul-gorevlisi", name: "Hasta Kabul Görevlisi", level: "Giriş", tasks: [
        ["Randevu Optimizasyonu", "Doktor takvimi ve gelmeme oranlarına göre randevu dağılımı önerir.", 4],
        ["Sigorta Provizyon Kontrolü", "Hasta sigorta bilgisinin ön kontrolünü yapar, eksikleri listeler.", 3],
        ["SMS Hatırlatma", "Yaklaşan randevular için hatırlatma mesajları hazırlar.", 2],
      ]},
      { id: "hasta-iliskileri-sorumlusu", name: "Hasta İlişkileri Sorumlusu", level: "Uzman", tasks: [
        ["Şikayet Sınıflandırma", "Hasta şikayetlerini birim ve konuya göre sınıflandırır.", 3],
        ["Memnuniyet Anketi Analizi", "Anket sonuçlarından birim bazlı memnuniyet raporu çıkarır.", 2],
      ]},
    ],
  },
  "tibbi-sekreterlik": {
    name: "Tıbbi Sekreterlik", short: "TSK",
    desc: "Epikriz, kodlama ve hastane raporları.",
    roles: [
      { id: "tibbi-sekreter", name: "Tıbbi Sekreter", level: "Uzman", tasks: [
        ["Epikriz Taslağı", "Dikte ve notlardan epikriz taslağı oluşturur (hekim onayı şart).", 6],
        ["ICD-10 Kod Önerisi", "Tanı metinlerinden ICD-10 kod önerileri üretir.", 4],
        ["Rapor Arşivleme", "Tetkik raporlarını hasta ve tarihe göre dosyalar.", 2],
      ]},
      { id: "hastane-yoneticisi", name: "Hastane Yöneticisi", level: "Yönetici", tasks: [
        ["Yatak Doluluk Analizi", "Servis bazında doluluk ve kalış süresi raporu üretir.", 2],
        ["SUT Fatura Kontrolü", "Fatura kalemlerini SUT kurallarına göre ön kontrolden geçirir.", 5],
      ]},
    ],
  },
};

export const SECTORS = [
  { id: "banka",     name: "Bankacılık",           icon: "bank",    tagline: "Kredi, uyum ve şube operasyonları.",      depts: ["krediler", "uyum", "sube", "ik", "finans", "bt"] },
  { id: "sigorta",   name: "Sigorta",              icon: "shield",  tagline: "Hasardan aktüeryaya, poliçeden acenteye.", depts: ["hasar", "akturya", "police", "ik", "finans", "bt"] },
  { id: "uretim",    name: "Üretim",               icon: "factory", tagline: "Planlama, kalite ve bakım.",              depts: ["planlama", "kalite", "bakim", "ik", "finans", "bt"] },
  { id: "tekstil",   name: "Tekstil",              icon: "thread",  tagline: "Tasarımdan numuneye, iplikten sevkiyata.", depts: ["tasarim", "numune", "tedarik", "ik", "finans", "bt"] },
  { id: "perakende", name: "Perakende & E-ticaret", icon: "bag",    tagline: "Mağaza, pazaryeri ve talep planlama.",    depts: ["eticaret", "magaza", "talep", "ik", "finans", "bt"] },
  { id: "saglik",    name: "Sağlık",               icon: "pulse",   tagline: "Hasta kabul, tıbbi sekreterlik, yönetim.", depts: ["hasta-kabul", "tibbi-sekreterlik", "ik", "finans", "bt"] },
];

export const SOON = [
  { name: "Lojistik", icon: "truck" },
  { name: "Hukuk", icon: "scale" },
  { name: "Enerji", icon: "bolt" },
  { name: "Eğitim", icon: "book" },
];

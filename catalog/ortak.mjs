/* =========================================================================
   Ortak departmanlar — birden çok sektörde aynı biçimde bulunan birimler.
   Görev formatı: [ad, açıklama, haftalık tahmini kazanç (saat)]
   Başka bir rolde aynı görev yalnızca adıyla ("Banka Mutabakatı") anılır.
   Kaynaklar: catalog/KAYNAKLAR.md
   ========================================================================= */

export const ORTAK = {
  /* ------------------------------ İNSAN KAYNAKLARI ------------------------------ */
  ik: {
    name: "İnsan Kaynakları", short: "İK", shared: true,
    desc: "İşe alım, bordro ve özlük, eğitim, performans ve ücret yönetimi.",
    roles: [
      { id: "ik-uzman-yardimcisi", name: "İnsan Kaynakları Uzman Yardımcısı", level: "Giriş", tasks: [
        ["CV Ön Eleme", "Gelen CV'leri ilanın zorunlu kriterlerine (deneyim yılı, eğitim, lokasyon, ehliyet vb.) göre ilk filtreden geçirir, elenenleri gerekçesiyle listeler.", 6],
        ["Mülakat Takvimi Planlama", "Aday ve mülakatçı müsaitliklerini eşleştirip çakışmasız mülakat takvimi ve davet listesi çıkarır.", 3],
        ["Aday Bilgilendirme E-postaları", "Süreç aşamasına (davet, olumsuz dönüş, teklif) göre kişiselleştirilmiş aday e-postaları hazırlar.", 2],
        ["Personel Özlük Dosyası Eksik Evrak Takibi", "Çalışan listesini zorunlu özlük evrakı listesiyle (kimlik, ikametgâh, diploma, sağlık raporu, adli sicil vb.) karşılaştırıp eksikleri çalışan bazında raporlar.", 3],
      ]},
      { id: "ise-alim-uzmani", name: "İşe Alım Uzmanı", level: "Uzman", tasks: [
        ["CV Raporlama", "Yüzlerce CV'yi okuyup deneyim, eğitim, dil ve yetkinliklere göre tek bir Excel raporunda toplar; ilana uyum yüzdesi hesaplar.", 8],
        ["İş İlanı Metni Hazırlama", "Pozisyonun görev ve niteliklerinden ayrımcı ifade içermeyen, yayına hazır iş ilanı metni hazırlar.", 2],
        ["Mülakat Soru Seti Hazırlama", "Pozisyonun yetkinliklerine göre yetkinlik bazlı mülakat soruları ve değerlendirme ölçütleri hazırlar.", 2],
        ["Mülakat Değerlendirme Özeti", "Birden fazla mülakatçının notlarını ve puanlarını aday başına tek sayfalık karşılaştırmalı özete çevirir.", 3],
        ["İşe Alım Süreç Raporu", "Açık pozisyonların aşama, bekleme süresi ve kaynak kanal kırılımıyla işe alım süresi (time-to-hire) raporunu üretir.", 2],
      ]},
      { id: "bordro-ve-ozluk-isleri-uzmani", name: "Bordro ve Özlük İşleri Uzmanı", level: "Uzman", tasks: [
        ["Brüt-Net Maaş Hesaplama", "Güncel yıl parametreleriyle (SGK, işsizlik, gelir vergisi dilimleri, damga vergisi, asgari ücret istisnası) brütten nete ve netten brüte maaş hesaplar.", 6],
        ["Puantaj Kontrolü", "Aylık puantajı (çalışılan gün, izin, rapor, fazla mesai) bordro öncesi tutarlılık kurallarına göre kontrol eder, hatalı satırları listeler.", 4],
        ["Fazla Mesai Hesaplama", "Haftalık 45 saati aşan çalışmaları ve tatil çalışmalarını 4857 sayılı Kanun'a göre hesaplayıp ücret karşılığını çıkarır.", 3],
        ["Yıllık İzin Hakedişi Hesaplama", "İşe giriş tarihi ve yaşa göre 4857 sayılı Kanun md. 53 kurallarıyla yıllık izin hakkını, kullanılanı ve kalanı hesaplar.", 3],
        ["Kıdem ve İhbar Tazminatı Hesaplama", "Giriş-çıkış tarihi ve giydirilmiş brüt ücretten kıdem tavanını dikkate alarak kıdem ve ihbar tazminatını, vergi ve damga kesintileriyle hesaplar.", 3],
        ["SGK Giriş-Çıkış Bildirge Kontrolü", "İşe giriş ve işten çıkış listesini SGK bildirge kayıtlarıyla karşılaştırıp yasal süresinde yapılmamış veya eksik bildirgeleri listeler.", 2],
      ]},
      { id: "egitim-ve-gelisim-uzmani", name: "Eğitim ve Gelişim Uzmanı", level: "Uzman", tasks: [
        ["Eğitim İhtiyaç Analizi", "Yönetici ve çalışan anketlerini, performans sonuçlarını birleştirip departman bazında eğitim ihtiyacı önceliklendirmesi çıkarır.", 3],
        ["Yıllık Eğitim Planı", "Eğitim ihtiyaçları, bütçe ve takvime göre yıllık eğitim planı ve katılımcı listelerini oluşturur.", 2],
        ["Eğitim Değerlendirme Raporu", "Katılım, sınav ve memnuniyet anketlerinden eğitim etkinlik raporu üretir.", 2],
      ]},
      { id: "ucretlendirme-ve-yan-haklar-uzmani", name: "Ücretlendirme ve Yan Haklar Uzmanı", level: "Uzman", tasks: [
        ["Ücret Bandı Karşılaştırması", "İç ücret verisini kademe/pozisyon bazında ücret araştırması bantlarıyla kıyaslar, bant dışı çalışanları işaretler.", 3],
        ["Zam Bütçesi Simülasyonu", "Farklı zam oranı ve performans matrisi senaryolarında toplam personel maliyetini ve kişi bazlı yeni ücretleri hesaplar.", 3],
        ["Performans Primi Hesaplama", "Hedef gerçekleşme ve prim skalasına göre çalışan bazında performans primi tutarlarını hesaplar.", 3],
      ]},
      { id: "insan-kaynaklari-muduru", name: "İnsan Kaynakları Müdürü", level: "Yönetici", tasks: [
        ["Personel Devir Oranı Analizi", "Giriş-çıkış verisinden departman, kıdem ve yönetici bazında devir oranı ve erken ayrılma sinyallerini çıkarır.", 3],
        ["Performans Değerlendirme Özeti", "Yönetici değerlendirmelerini ve hedef sonuçlarını çalışan başına tek sayfalık özete ve dağılım raporuna çevirir.", 5],
        ["Çalışan Bağlılık Anketi Analizi", "Anket puanlarını ve açık uçlu yorumları birim bazında analiz edip temel sorun başlıklarını özetler.", 3],
        ["Personel Maliyet Bütçesi", "Kadro planı, ücretler ve yan haklardan aylık personel maliyet bütçesi ve gerçekleşme karşılaştırması üretir.", 3],
      ]},
    ],
  },

  /* ------------------------------ MUHASEBE ------------------------------ */
  muhasebe: {
    name: "Muhasebe", short: "MHS", shared: true,
    desc: "Fatura, cari, banka, mutabakat, vergi ve ay sonu kapanış işleri.",
    roles: [
      { id: "on-muhasebe-elemani", name: "Ön Muhasebe Elemanı", level: "Giriş", tasks: [
        ["e-Fatura Okuma ve Listeleme", "e-Fatura / e-Arşiv UBL-TR XML dosyalarını okuyup fatura no, tarih, VKN, matrah, KDV oranı kırılımı ve tevkifat bilgisiyle Excel listesine çevirir.", 8],
        ["Cari Hesap Ekstresi Hazırlama", "Muhasebe hareketlerinden müşteri/tedarikçi bazında dönem ekstresi ve bakiye özeti hazırlar.", 3],
        ["Tahsilat ve Ödeme Listesi", "Vadesi gelen alacak ve borçlardan haftalık tahsilat ve ödeme listesini çıkarır.", 2],
        ["VKN / TCKN / IBAN Doğrulama", "Cari kart listesindeki vergi kimlik no, TC kimlik no ve IBAN'ları algoritmik olarak doğrular, hatalı ve mükerrer kayıtları listeler.", 2],
      ]},
      { id: "muhasebe-elemani", name: "Muhasebe Elemanı", level: "Giriş", tasks: [
        ["Banka Mutabakatı", "Banka ekstresi ile muhasebe defterini (102 Bankalar) eşleştirir; eşleşenleri, açık kalemleri ve olası nedenlerini raporlar.", 6],
        ["Cari Hesap Mutabakatı", "Firma ile karşı tarafın cari ekstrelerini karşılaştırıp fatura, ödeme ve iade bazında farkları ve olası nedenlerini çıkarır.", 5],
        ["Form Ba-Bs Kontrolü", "Alış ve satış faturalarından Ba-Bs bildirim tutarlarını VKN bazında hesaplar, karşı tarafın bildirdikleriyle uyumsuzlukları listeler.", 4],
        ["Masraf Fişi Kategorileme", "Çalışan masraf fişlerini ve kredi kartı harcamalarını hesap planına göre sınıflandırıp muhasebe kayıt önerisi çıkarır.", 3],
        ["Fatura KDV Tutarlılık Kontrolü", "Fatura satırlarında matrah, KDV oranı ve KDV tutarı tutarlılığını, tevkifat oranlarını ve yuvarlama hatalarını kontrol eder.", 3],
      ]},
      { id: "genel-muhasebe-uzmani", name: "Genel Muhasebe Uzmanı", level: "Uzman", tasks: [
        ["Mizan Kontrolü", "Aylık mizanda ters bakiye veren hesapları, mantık dışı bakiyeleri ve önceki aya göre olağandışı değişimleri işaretler.", 4],
        ["Ay Sonu Kapanış Kontrol Listesi", "Kapanış adımlarını (tahakkuklar, amortisman, kur değerleme, mutabakatlar) sorumlu ve durumla takip eden kontrol listesi üretir.", 2],
        ["Dönem Sonu Kur Değerleme", "Döviz cinsinden hesapların bakiyelerini dönem sonu kurlarıyla değerleyip kur farkı kayıtlarını hesaplar.", 3],
        ["Amortisman Hesaplama", "Demirbaş listesinden faydalı ömür ve yönteme (normal/azalan bakiyeler) göre dönem amortismanını ve birikmiş amortismanı hesaplar.", 2],
        ["KDV Beyannamesi Ön Kontrolü", "Alış-satış listeleri ve mizanla KDV beyannamesi tutarlarını karşılaştırıp uyumsuzlukları beyan öncesi yakalar.", 3],
      ]},
      { id: "maliyet-muhasebesi-uzmani", name: "Maliyet Muhasebesi Uzmanı", level: "Uzman", tasks: [
        ["Ürün Maliyeti Hesaplama", "Reçete (ürün ağacı), hammadde birim maliyeti, işçilik ve genel üretim giderlerinden ürün birim maliyetini hesaplar.", 5],
        ["Standart-Fiili Maliyet Sapma Analizi", "Standart maliyet ile fiili maliyet arasındaki farkı miktar ve fiyat sapması olarak ayrıştırır.", 3],
        ["Stok Değerleme", "Stok hareketlerinden ağırlıklı ortalama veya FIFO yöntemiyle dönem sonu stok değerini ve satılan malın maliyetini hesaplar.", 3],
      ]},
      { id: "muhasebe-muduru", name: "Muhasebe Müdürü", level: "Yönetici", tasks: [
        ["Mizandan Mali Tablo Hazırlama", "Tekdüzen hesap planına göre mizandan bilanço ve gelir tablosu üretir.", 4],
        ["Vergi Takvimi Takibi", "Beyanname ve ödeme tarihlerini şirketin yükümlülüklerine göre listeleyip yaklaşan son tarihleri hatırlatır.", 1],
        ["Denetim Hazırlık Dosyası", "Bağımsız denetim talep listesine göre belge ve mutabakatların hazır olup olmadığını takip eden dosya hazırlar.", 2],
      ]},
    ],
  },

  /* ------------------------------ FİNANS ------------------------------ */
  finans: {
    name: "Finans", short: "FİN", shared: true,
    desc: "Nakit yönetimi, bütçe ve raporlama, finansal analiz.",
    roles: [
      { id: "finans-uzmani", name: "Finans Uzmanı", level: "Uzman", tasks: [
        ["Nakit Akış Tahmini", "Vadeli alacak, borç, çek-senet ve sabit ödemelerden haftalık (13 hafta) nakit akış tahmini ve nakit açığı uyarısı üretir.", 4],
        ["Alacak Yaşlandırma", "Açık faturaları vade gününe göre (0-30, 31-60, 61-90, 90+) müşteri bazında yaşlandırır, riskli alacakları işaretler.", 3],
        ["Çek-Senet Portföy Takibi", "Alınan ve verilen çek-senetleri vade, banka ve durum bazında takip eder, haftalık vade listesi çıkarır.", 3],
        ["Banka Limit ve Risk Raporu", "Banka bazında kredi limitleri, kullanımlar, teminatlar ve faiz oranlarını tek tabloda toplar.", 2],
        ["Ödeme Önceliklendirme", "Nakit durumuna göre ödeme listesini vade, gecikme cezası ve tedarikçi önemine göre önceliklendirir.", 2],
      ]},
      { id: "butce-ve-raporlama-uzmani", name: "Bütçe ve Raporlama Uzmanı", level: "Uzman", tasks: [
        ["Bütçe-Gerçekleşen Sapma Raporu", "Bütçe ile gerçekleşeni masraf merkezi ve hesap bazında karşılaştırır, önemli sapmaları açıklama alanıyla raporlar.", 5],
        ["Yönetim Raporu (KPI)", "Aylık yönetim raporunun tablo ve grafiklerini muhasebe ve satış verisinden otomatik hazırlar.", 6],
        ["Departman Bütçelerinin Konsolidasyonu", "Departmanlardan gelen bütçe şablonlarını tek dosyada birleştirir, format ve toplam hatalarını kontrol eder.", 3],
      ]},
      { id: "finansal-analist", name: "Finansal Analist", level: "Uzman", tasks: [
        ["Mali Tablo Rasyo Analizi", "Bilanço ve gelir tablosundan likidite, kaldıraç, kârlılık ve faaliyet rasyolarını dönemsel karşılaştırmalı hesaplar.", 4],
        ["Yatırım Fizibilitesi (NPV/IRR)", "Yatırım nakit akışlarından net bugünkü değer, iç verim oranı ve geri dönüş süresini hesaplar.", 3],
        ["Senaryo ve Duyarlılık Analizi", "Kur, faiz, satış ve maliyet varsayımlarına göre kârlılık ve nakit senaryoları üretir.", 3],
      ]},
      { id: "finans-muduru", name: "Finans Müdürü", level: "Yönetici", tasks: [
        ["Kur Riski Pozisyon Raporu", "Döviz cinsinden varlık ve yükümlülüklerden döviz pozisyonunu ve kur şoku senaryolarındaki etkisini hesaplar.", 2],
        ["Konsolide Finansal Rapor", "Grup şirketlerinin mizanlarını birleştirip karşılıklı işlemleri eleyerek konsolide tabloları üretir.", 5],
        ["Yönetim Kurulu Finans Sunumu Özeti", "Dönem finansal sonuçlarından yönetim kurulu için kısa, yorumlu özet metni hazırlar.", 2],
      ]},
    ],
  },

  /* ------------------------------ BİLGİ TEKNOLOJİLERİ ------------------------------ */
  bt: {
    name: "Bilgi Teknolojileri", short: "BT", shared: true,
    desc: "Kullanıcı desteği, sistemler, yazılım, bilgi güvenliği ve KVKK.",
    roles: [
      { id: "bt-destek-uzmani", name: "BT Destek Uzmanı", level: "Giriş", tasks: [
        ["Destek Talebi Sınıflandırma", "Gelen destek taleplerini (e-posta / ticket) kategori, öncelik ve ilgili ekibe göre otomatik etiketler.", 5],
        ["Sık Sorulan Sorulara Cevap Taslağı", "Kullanıcı sorularına şirket bilgi bankasından cevap taslağı üretir.", 5],
        ["Donanım Zimmet Envanteri", "Bilgisayar, telefon ve çevre birimlerinin zimmet, garanti ve yenileme tarihlerini tek envanterde takip eder.", 2],
      ]},
      { id: "sistem-uzmani", name: "Sistem Uzmanı", level: "Uzman", tasks: [
        ["Log Anomali Tespiti", "Sunucu ve uygulama loglarında olağandışı hata, oturum ve erişim örüntülerini bulup özetler.", 4],
        ["Yedekleme Doğrulama Raporu", "Yedekleme yazılımı çıktılarından başarısız, eksik veya geciken yedekleri raporlar.", 2],
        ["SSL Sertifika ve Lisans Süre Takibi", "Sertifika, alan adı ve yazılım lisanslarının bitiş tarihlerini izleyip yaklaşanları listeler.", 1],
        ["Kapasite Raporu", "Disk, bellek ve işlemci kullanım verisinden büyüme eğilimi ve dolma tarihi tahmini üretir.", 2],
      ]},
      { id: "bilgi-guvenligi-uzmani", name: "Bilgi Güvenliği Uzmanı", level: "Uzman", tasks: [
        ["Zafiyet Tarama Önceliklendirme", "Zafiyet tarama çıktısını varlık kritikliği ve CVSS skoruna göre önceliklendirip aksiyon listesi çıkarır.", 4],
        ["Kullanıcı Yetki Gözden Geçirme", "Sistemlerdeki kullanıcı-yetki listelerini İK aktif personel listesiyle karşılaştırıp ayrılmış personeli ve aşırı yetkileri bulur.", 3],
        ["KVKK Kişisel Veri İşleme Envanteri", "Departmanlardan toplanan veri işleme bilgilerinden kişisel veri işleme envanteri tablosu oluşturur.", 3],
        ["Oltalama E-postası Analizi", "Şüpheli e-postanın başlık, bağlantı ve içerik işaretlerini inceleyip risk değerlendirmesi yazar.", 2],
      ]},
      { id: "bt-muduru", name: "BT Müdürü", level: "Yönetici", tasks: [
        ["BT Maliyet Raporu", "Lisans, bulut, donanım ve hizmet harcamalarını birim ve kalem bazında raporlar.", 3],
        ["BT Proje Portföy Durumu", "Projelerin bütçe, takvim ve risk durumunu tek bir yönetim tablosunda toplar.", 2],
      ]},
    ],
  },

  /* ------------------------------ SATIN ALMA ------------------------------ */
  satinalma: {
    name: "Satın Alma", short: "SAT", shared: true,
    desc: "Talep, teklif, sipariş, tedarikçi yönetimi ve harcama analizi.",
    roles: [
      { id: "satin-alma-uzman-yardimcisi", name: "Satın Alma Uzman Yardımcısı", level: "Giriş", tasks: [
        ["Satın Alma Talebi Konsolidasyonu", "Departmanlardan gelen satın alma taleplerini malzeme ve tedarikçi bazında birleştirip teklif istenecek listeyi çıkarır.", 3],
        ["Sipariş Termin Takibi", "Açık siparişleri termin tarihine göre izler, geciken ve gecikme riski olanları tedarikçi bazında listeler.", 4],
      ]},
      { id: "satin-alma-uzmani", name: "Satın Alma Uzmanı", level: "Uzman", tasks: [
        ["Teklif Karşılaştırma", "Tedarikçi tekliflerini fiyat, vade, termin, kalite ve garanti kriterleriyle ağırlıklı puanlayarak karşılaştırma tablosu üretir.", 4],
        ["Tedarikçi Performans Değerlendirme", "Teslim zamanı, kalite ret oranı ve fiyat uyumundan tedarikçi karnesi hazırlar.", 3],
        ["Satın Alma Fiyat Geçmişi Analizi", "Malzeme bazında geçmiş alım fiyatlarını kur ve enflasyonla birlikte analiz edip olağandışı fiyat artışlarını işaretler.", 2],
      ]},
      { id: "satin-alma-muduru", name: "Satın Alma Müdürü", level: "Yönetici", tasks: [
        ["Harcama Analizi", "Satın alma verisini kategori ve tedarikçi bazında ABC sınıflandırmasıyla analiz edip pazarlık önceliklerini çıkarır.", 3],
        ["Tedarikçi Risk Skoru", "Tek kaynak bağımlılığı, finansal durum, gecikme ve kalite verisinden tedarikçi risk skoru üretir.", 2],
        ["Satın Alma Tasarruf Raporu", "Pazarlık ve alternatif tedarik kaynaklı tasarrufları bütçe fiyatına göre raporlar.", 2],
      ]},
    ],
  },

  /* ------------------------------ HUKUK ------------------------------ */
  hukuk: {
    name: "Hukuk", short: "HUK", shared: true,
    desc: "Sözleşmeler, dava ve icra takibi, KVKK ve mevzuat uyumu.",
    roles: [
      { id: "avukat", name: "Avukat", level: "Uzman", tasks: [
        ["Sözleşme Ön İnceleme", "Sözleşme metnindeki riskli maddeleri (ceza şartı, fesih, sorumluluk sınırı, yetkili mahkeme, gizlilik) işaretleyip öneriyle özetler.", 6],
        ["İhtarname Taslağı", "Olay özeti ve talepten noter ihtarnamesi taslağı hazırlar.", 2],
        ["Dava ve Duruşma Takvimi", "Dava listesinden duruşma tarihleri ve kesin süreleri takip eden takvim ve hatırlatma listesi çıkarır.", 3],
        ["İcra Takip Durum Raporu", "İcra dosyalarının aşama, tahsilat ve masraf bilgilerini tek raporda toplar.", 3],
        ["KVKK Aydınlatma Metni Taslağı", "Veri işleme amaçlarından KVKK'ya uygun aydınlatma metni taslağı hazırlar.", 2],
      ]},
      { id: "hukuk-muduru", name: "Hukuk Müdürü", level: "Yönetici", tasks: [
        ["Dava Risk ve Karşılık Raporu", "Devam eden davaların tutar ve kaybetme olasılığına göre risk ve karşılık tablosunu hazırlar.", 2],
        ["Sözleşme Envanteri ve Süre Takibi", "Sözleşmelerin taraf, bitiş, yenileme ve fesih bildirim tarihlerini izleyen envanter oluşturur.", 2],
        ["Mevzuat Değişikliği Özeti", "Resmî Gazete'de yayımlanan düzenlemeleri şirketin faaliyet alanına göre süzer ve etkisini özetler.", 3],
      ]},
    ],
  },

  /* ------------------------------ İDARİ İŞLER ------------------------------ */
  idari: {
    name: "İdari İşler", short: "İDR", shared: true,
    desc: "Araç filosu, demirbaş, hizmet alımları ve ofis yönetimi.",
    roles: [
      { id: "idari-isler-sorumlusu", name: "İdari İşler Sorumlusu", level: "Uzman", tasks: [
        ["Araç Filosu Takibi", "Şirket araçlarının muayene, trafik sigortası, kasko, bakım ve lastik değişim tarihlerini takip edip yaklaşanları listeler.", 2],
        ["Demirbaş ve Zimmet Takibi", "Demirbaşların konum, zimmet ve sayım durumunu takip eder, sayım farklarını raporlar.", 2],
        ["Hizmet Alımı Fatura Kontrolü", "Yemek, servis, temizlik ve güvenlik faturalarını sözleşme birim fiyatı ve fiilî kullanım (kişi/gün, sefer) ile karşılaştırır.", 3],
        ["Abonelik ve Sözleşme Takibi", "Kira, elektrik, su, internet ve bakım sözleşmelerinin tutar, vade ve yenileme tarihlerini izler.", 1],
      ]},
    ],
  },

  /* ------------------------------ PAZARLAMA ------------------------------ */
  pazarlama: {
    name: "Pazarlama", short: "PZR", shared: true,
    desc: "Marka, dijital pazarlama, kampanya ve müşteri analitiği.",
    roles: [
      { id: "dijital-pazarlama-uzmani", name: "Dijital Pazarlama Uzmanı", level: "Uzman", tasks: [
        ["Reklam Kampanyası Performans Raporu", "Reklam platformu dışa aktarımlarından harcama, tıklama, dönüşüm ve ROAS'ı kampanya bazında birleştirip raporlar.", 4],
        ["Sosyal Medya İçerik Takvimi", "Kampanya ve özel günlere göre aylık sosyal medya içerik takvimi ve gönderi metni taslakları hazırlar.", 3],
        ["E-bülten Metni Hazırlama", "Kampanya bilgisinden konu satırı alternatifleriyle e-bülten metni hazırlar.", 2],
      ]},
      { id: "pazarlama-uzmani", name: "Pazarlama Uzmanı", level: "Uzman", tasks: [
        ["Müşteri Segmentasyonu (RFM)", "Satış geçmişinden müşterileri yenilik, sıklık ve tutara göre segmentlere ayırır, segment bazlı aksiyon önerir.", 3],
        ["Rakip Analizi Raporu", "Rakip fiyat, ürün ve kampanya bilgilerini derleyip karşılaştırmalı rapor hazırlar.", 3],
        ["Anket Sonuç Analizi", "Pazar araştırması anketlerinin kapalı ve açık uçlu cevaplarını analiz edip bulguları özetler.", 3],
      ]},
      { id: "pazarlama-muduru", name: "Pazarlama Müdürü", level: "Yönetici", tasks: [
        ["Pazarlama Bütçesi Takibi", "Kanal ve kampanya bazında pazarlama bütçesi, harcama ve getirisini takip eder.", 2],
        ["Basın Bülteni Taslağı", "Duyuru bilgisinden kurumsal dilde basın bülteni taslağı hazırlar.", 1],
      ]},
    ],
  },

  /* ------------------------------ SATIŞ ------------------------------ */
  satis: {
    name: "Satış", short: "STŞ", shared: true,
    desc: "Saha satış, satış destek ve satış yönetimi.",
    roles: [
      { id: "satis-temsilcisi", name: "Satış Temsilcisi", level: "Giriş", tasks: [
        ["Müşteri Ziyaret Planı", "Müşteri listesi, adres ve ziyaret sıklığına göre haftalık ziyaret planı ve rota sırası çıkarır.", 3],
        ["Teklif Hazırlama", "Fiyat listesi, iskonto kuralları ve ürün seçiminden müşteriye gönderilecek teklif dosyasını hazırlar.", 3],
        ["Ziyaret Notu Özeti", "Müşteri görüşme notlarını CRM'e girilecek özet ve takip aksiyonlarına çevirir.", 2],
      ]},
      { id: "satis-destek-uzmani", name: "Satış Destek Uzmanı", level: "Uzman", tasks: [
        ["Satış Hedef-Gerçekleşme Raporu", "Temsilci, bölge ve ürün grubu bazında hedef-gerçekleşme ve geçen yıl karşılaştırma raporu üretir.", 4],
        ["Teklif Takip Listesi", "Verilen tekliflerin durumunu, bekleme süresini ve dönüşüm oranını takip eder.", 2],
        ["Müşteri Risk ve Limit Takibi", "Açık bakiye, vadesi geçmiş alacak ve kredi limitine göre sevkiyat öncesi riskli müşterileri işaretler.", 2],
        ["Fiyat Listesi Güncelleme", "Maliyet değişimi ve kâr marjı kurallarına göre yeni fiyat listesini ve değişim raporunu hazırlar.", 2],
      ]},
      { id: "satis-muduru", name: "Satış Müdürü", level: "Yönetici", tasks: [
        ["Satış Tahmini", "Geçmiş satışlar ve mevsimsellikten ürün grubu ve bölge bazında aylık satış tahmini üretir.", 3],
        ["Satış Primi Hesaplama", "Prim sistemine göre temsilci bazında hedef gerçekleşme ve prim tutarlarını hesaplar.", 2],
      ]},
    ],
  },

  /* ------------------------------ MÜŞTERİ HİZMETLERİ ------------------------------ */
  "musteri-hizmetleri": {
    name: "Müşteri Hizmetleri ve Çağrı Merkezi", short: "MÜŞ", shared: true,
    desc: "Müşteri talepleri, şikâyetler, çağrı merkezi performansı.",
    roles: [
      { id: "musteri-temsilcisi", name: "Müşteri Temsilcisi", level: "Giriş", tasks: [
        ["Müşteri Talebi Sınıflandırma ve Cevap Taslağı", "Gelen e-posta, form ve mesajları konu ve aciliyete göre sınıflandırır, şirket bilgilerine dayanarak cevap taslağı hazırlar.", 8],
        ["Görüşme Kaydı Özeti", "Müşteri görüşme notlarını talep, yapılan işlem ve bekleyen aksiyon olarak standart kayıt formatına çevirir.", 3],
      ]},
      { id: "musteri-deneyimi-uzmani", name: "Müşteri Deneyimi Uzmanı", level: "Uzman", tasks: [
        ["Şikâyet Analizi", "Şikâyet kayıtlarını konu, ürün ve kök neden başlıklarına kümeleyip eğilim raporu çıkarır.", 4],
        ["Memnuniyet (NPS) Anketi Analizi", "NPS ve memnuniyet anketlerinin puan ve yorumlarını analiz edip iyileştirme önerileri çıkarır.", 3],
        ["Çağrı Kalite Değerlendirmesi", "Görüşme dökümünü kalite formundaki kriterlere göre puanlayıp geri bildirim notu hazırlar.", 3],
      ]},
      { id: "cagri-merkezi-yoneticisi", name: "Çağrı Merkezi Yöneticisi", level: "Yönetici", tasks: [
        ["Çağrı Merkezi Performans Raporu", "Çağrı verisinden karşılama oranı, servis seviyesi, ortalama görüşme süresi ve ilk temasta çözüm oranını hesaplar.", 3],
        ["Temsilci İhtiyacı Hesaplama (Erlang C)", "Saatlik çağrı hacmi ve hedef servis seviyesine göre Erlang C ile gereken temsilci sayısını ve vardiya ihtiyacını hesaplar.", 3],
      ]},
    ],
  },

  /* ------------------------------ İŞ SAĞLIĞI VE GÜVENLİĞİ ------------------------------ */
  isg: {
    name: "İş Sağlığı ve Güvenliği", short: "İSG", shared: true,
    desc: "Risk değerlendirmesi, eğitim ve muayene takibi, kaza kayıtları.",
    roles: [
      { id: "is-guvenligi-uzmani", name: "İş Güvenliği Uzmanı", level: "Uzman", tasks: [
        ["Risk Değerlendirmesi (Fine-Kinney)", "Tehlike listesinde olasılık, frekans ve şiddet puanlarından Fine-Kinney risk skorunu ve risk sınıfını hesaplar, önlem önceliklerini sıralar.", 4],
        ["İSG Eğitim ve Muayene Takibi", "Çalışanların tehlike sınıfına göre periyodik İSG eğitimi ve sağlık muayenesi tarihlerini takip eder, süresi geçenleri listeler.", 3],
        ["Periyodik Kontrol Takibi", "Kaldırma ekipmanı, basınçlı kap, elektrik tesisatı ve yangın ekipmanının periyodik kontrol tarihlerini izler.", 2],
        ["İş Kazası ve Ramak Kala Analizi", "Kaza ve ramak kala kayıtlarından bölüm, neden ve zaman bazlı analiz ve kaza sıklık oranlarını üretir.", 2],
        ["Acil Durum Eylem Planı Taslağı", "İşyeri bilgilerinden acil durum senaryoları, ekipler ve tahliye adımlarını içeren plan taslağı hazırlar.", 2],
      ]},
    ],
  },

  /* ------------------------------ DIŞ TİCARET ------------------------------ */
  "dis-ticaret": {
    name: "Dış Ticaret", short: "DTC", shared: true,
    desc: "İhracat ve ithalat operasyonu, evrak, gümrük ve teşvik takibi.",
    roles: [
      { id: "dis-ticaret-uzmani", name: "Dış Ticaret Uzmanı", level: "Uzman", tasks: [
        ["Proforma Fatura Hazırlama", "Sipariş, fiyat ve teslim şekli (Incoterms) bilgilerinden proforma fatura ve çeki listesi hazırlar.", 3],
        ["İhracat Evrak Kontrol Listesi", "Ülke, teslim şekli ve ödeme yöntemine göre gereken evrakları listeler; fatura, çeki listesi ve menşe bilgilerinin tutarlılığını kontrol eder.", 3],
        ["GTİP Sınıflandırma Önerisi", "Ürün tanımından olası GTİP kodlarını gerekçesiyle önerir (gümrük müşaviri onayı gerekir).", 2],
        ["Akreditif Evrak Uygunluk Kontrolü", "Akreditif şartlarını ibraz edilecek evraklarla karşılaştırıp olası rezervleri (uyumsuzlukları) işaretler.", 3],
        ["Dahilde İşleme İzin Belgesi Takibi", "DİİB kapsamındaki ithalat ve ihracat taahhütlerini, kalan miktarları ve süre sonunu takip eder.", 2],
      ]},
    ],
  },

  /* ------------------------------ LOJİSTİK VE DEPO (işletme içi) ------------------------------ */
  depo: {
    name: "Lojistik ve Depo", short: "DPO", shared: true,
    desc: "Depo stokları, sayım, sevkiyat planlama ve nakliye maliyetleri.",
    roles: [
      { id: "depo-sorumlusu", name: "Depo Sorumlusu", level: "Uzman", tasks: [
        ["Stok Sayım Fark Analizi", "Sayım sonuçlarını sistem stoğuyla karşılaştırıp ürün ve lokasyon bazında farkları, tutar etkisini ve olası nedenleri raporlar.", 4],
        ["Stok Yaşlandırma Raporu", "Son hareket tarihine göre hareketsiz ve yavaş dönen stokları tutar bazında yaşlandırır.", 2],
        ["Depo Yerleşim (ABC) Önerisi", "Çıkış sıklığına göre ürünleri ABC sınıflandırıp sık çıkanların sevkiyata yakın konumlanmasını önerir.", 2],
      ]},
      { id: "lojistik-uzmani", name: "Lojistik Uzmanı", level: "Uzman", tasks: [
        ["Sevkiyat Araç Planlama", "Sipariş hacmi, ağırlık ve palet sayısına göre araç tipi ve doluluk planı çıkarır.", 3],
        ["Nakliye Maliyet Analizi", "Nakliye faturalarını sefer, ağırlık ve bölge bazında analiz edip birim taşıma maliyetini hesaplar.", 2],
        ["Teslimat Performansı (OTIF) Raporu", "Siparişlerin zamanında ve eksiksiz teslim oranını müşteri ve taşıyıcı bazında raporlar.", 2],
      ]},
    ],
  },
};

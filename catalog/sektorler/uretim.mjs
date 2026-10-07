/* Üretim / Sanayi (genel imalat) — fabrika bölümleri eleman.net fabrika rehberi ve
   "Üretim Planlama Mühendisi", "Fabrika Bakım Yöneticisi" gibi ilanlara dayanır. */

export const URETIM = {
  "uretim-uretim": {
    name: "Üretim", short: "ÜRT",
    desc: "Hat ve vardiya yönetimi, verimlilik, fire ve iş emirleri.",
    roles: [
      { id: "vardiya-amiri", name: "Vardiya Amiri", level: "Uzman", tasks: [
        ["Vardiya Üretim Raporu", "Vardiya bazında üretim adedi, hedef, duruşlar ve fire bilgisinden vardiya raporu üretir.", 3],
        ["Vardiya Devir Notu Özeti", "Vardiya defteri notlarını açık sorunlar, arızalar ve sonraki vardiyaya aksiyonlar olarak özetler.", 2],
      ]},
      { id: "uretim-muhendisi", name: "Üretim Mühendisi", level: "Uzman", tasks: [
        ["OEE Hesaplama", "Planlı süre, duruşlar, çevrim süresi ve hatalı adetlerden makine/hat bazında kullanılabilirlik, performans, kalite ve OEE'yi hesaplar.", 3],
        ["Duruş Nedeni Pareto Analizi", "Makine duruş kayıtlarını neden ve süreye göre Pareto analizine tabi tutar.", 2],
        ["Fire ve Verim Analizi", "Hammadde girişi ile mamul çıkışından ürün ve hat bazında fire oranını ve verim kaybı maliyetini hesaplar.", 3],
      ]},
      { id: "uretim-muduru", name: "Üretim Müdürü", level: "Yönetici", tasks: [
        ["Üretim Performans Panosu", "OEE, plan uyumu, fire ve iş kazası göstergelerini hat ve hafta bazında tek panoda toplar.", 2],
      ]},
    ],
  },

  "uretim-planlama": {
    name: "Üretim Planlama", short: "PLN",
    desc: "Çizelgeleme, malzeme ihtiyaç planlama ve kapasite.",
    roles: [
      { id: "uretim-planlama-uzmani", name: "Üretim Planlama Uzmanı", level: "Uzman", tasks: [
        ["Haftalık Üretim Çizelgesi", "Sipariş, termin, rota ve makine kapasitesinden makine bazlı haftalık üretim çizelgesi oluşturur.", 8],
        ["Malzeme İhtiyaç Planlaması (MRP)", "Ürün ağacı, siparişler ve mevcut stoktan net malzeme ihtiyacını ve sipariş verilmesi gereken tarihleri hesaplar.", 6],
        ["Kapasite Darboğaz Analizi", "İş merkezi bazında yükleme oranlarını hesaplayıp darboğaz haftalarını ve kaydırma önerilerini çıkarır.", 3],
        ["Plan Uyum Raporu", "Planlanan ile gerçekleşen üretimi iş emri bazında karşılaştırıp sapma nedenlerini raporlar.", 2],
      ]},
      { id: "stok-planlama-uzmani", name: "Stok Planlama Uzmanı", level: "Uzman", tasks: [
        ["Yeniden Sipariş Noktası Hesabı", "Tüketim dağılımı, tedarik süresi ve hizmet düzeyinden emniyet stoğu ve yeniden sipariş noktasını hesaplar.", 3],
        ["Stok ABC-XYZ Analizi", "Stokları değer (ABC) ve talep değişkenliği (XYZ) açısından sınıflandırıp her sınıf için stok politikası önerir.", 3],
        "Stok Yaşlandırma Raporu",
      ]},
    ],
  },

  "uretim-kalite": {
    name: "Kalite Güvence ve Kalite Kontrol", short: "KLT",
    desc: "Giriş, proses ve final kontrol, istatistiksel proses kontrol, DÖF ve 8D.",
    roles: [
      { id: "kalite-kontrol-elemani", name: "Kalite Kontrol Elemanı", level: "Giriş", tasks: [
        ["AQL Örneklem Planı", "Parti büyüklüğü, muayene seviyesi ve AQL değerine göre ISO 2859-1 tablosundan örneklem sayısını ve kabul/ret sayılarını belirler.", 2],
        ["Ölçüm Verisi Uygunluk Kontrolü", "Ölçüm cihazı çıktılarını teknik resim toleranslarıyla karşılaştırıp uygunsuz ölçüleri işaretler.", 3],
      ]},
      { id: "kalite-muhendisi", name: "Kalite Mühendisi", level: "Uzman", tasks: [
        ["SPC Kontrol Grafiği ve Cp/Cpk", "Proses ölçümlerinden X̄-R kontrol grafiği, kontrol dışı noktalar ve Cp/Cpk yeterlilik indekslerini hesaplar.", 4],
        ["8D Rapor Taslağı", "Uygunsuzluk ve müşteri şikâyeti bilgilerinden 8D raporunun disiplin adımlarını içeren taslağı hazırlar.", 4],
        ["DÖF Takibi", "Düzeltici ve önleyici faaliyetleri sorumlu, termin ve etkinlik doğrulamasıyla takip eder.", 2],
        ["Müşteri Şikâyeti Kök Neden Analizi", "Şikâyet kayıtlarını ürün, hata türü ve üretim partisine göre kümeleyip olası kök nedenleri önerir.", 3],
      ]},
      { id: "kalite-muduru", name: "Kalite Müdürü", level: "Yönetici", tasks: [
        ["Kalitesizlik Maliyeti Raporu", "Hurda, yeniden işleme, iade ve garanti maliyetlerini ürün ve neden bazında raporlar.", 2],
        ["Tedarikçi Kalite Karnesi", "Giriş kalite kontrol verisinden tedarikçi bazında ret oranı (ppm) ve karne üretir.", 2],
        ["İç Tetkik Bulgu Takibi", "ISO 9001 iç tetkik bulgularını madde, birim ve kapanış durumuna göre takip eder.", 1],
      ]},
    ],
  },

  "uretim-bakim": {
    name: "Bakım Onarım", short: "BKM",
    desc: "Arıza, planlı ve kestirimci bakım, yedek parça.",
    roles: [
      { id: "bakim-teknisyeni", name: "Bakım Teknisyeni", level: "Giriş", tasks: [
        ["Arıza Kaydı Sınıflandırma", "Serbest metin arıza kayıtlarını makine, arıza tipi ve müdahale türüne göre etiketler.", 3],
        ["Periyodik Bakım İş Emri Listesi", "Bakım planından haftalık periyodik bakım iş emirlerini ve gereken malzemeleri çıkarır.", 2],
      ]},
      { id: "bakim-muhendisi", name: "Bakım Mühendisi", level: "Uzman", tasks: [
        ["MTBF ve MTTR Analizi", "Arıza kayıtlarından makine bazında arızalar arası ortalama süre ve ortalama onarım süresini hesaplar.", 3],
        ["Kestirimci Bakım Uyarıları", "Titreşim, sıcaklık ve akım sensör verisinde eşik ve eğilim analiziyle arıza riski yüksek ekipmanları belirler.", 4],
        ["Yedek Parça Kritik Stok Önerisi", "Arıza geçmişi ve tedarik süresine göre kritik yedek parçaların minimum stok seviyesini önerir.", 2],
      ]},
      { id: "bakim-muduru", name: "Bakım Müdürü", level: "Yönetici", tasks: [
        ["Bakım Maliyet ve Performans Raporu", "Bakım işçilik, malzeme ve dış hizmet maliyetlerini planlı/plansız bakım oranıyla raporlar.", 2],
      ]},
    ],
  },

  "uretim-endustri": {
    name: "Endüstri Mühendisliği ve Süreç Geliştirme", short: "END",
    desc: "İş etüdü, hat dengeleme, yalın üretim ve iyileştirme projeleri.",
    roles: [
      { id: "endustri-muhendisi", name: "Endüstri Mühendisi", level: "Uzman", tasks: [
        ["Zaman Etüdü ve Standart Süre", "Kronometre ölçümlerinden performans ve tolerans payıyla operasyon standart sürelerini hesaplar.", 3],
        ["Hat Dengeleme", "Operasyon süreleri ve öncelik ilişkilerinden hedef çevrim süresine göre istasyon atamasını ve hat verimini hesaplar.", 3],
        ["Kaizen Önerisi Değerlendirme", "Çalışan iyileştirme önerilerini kazanç, maliyet ve uygulanabilirliğe göre puanlar.", 2],
      ]},
    ],
  },

  "uretim-arge": {
    name: "Ar-Ge ve Ürün Geliştirme", short: "ARG",
    desc: "Ürün tasarımı, prototip, test ve teşvik projeleri.",
    roles: [
      { id: "ar-ge-muhendisi", name: "Ar-Ge Mühendisi", level: "Uzman", tasks: [
        ["Test Sonuçları Spesifikasyon Karşılaştırması", "Prototip test sonuçlarını spesifikasyon limitleriyle karşılaştırıp uygunsuzlukları ve eğilimleri raporlar.", 3],
        ["Ürün Ağacı (BOM) Karşılaştırma", "İki ürün ağacı versiyonunu karşılaştırıp eklenen, çıkarılan ve miktarı değişen kalemleri listeler.", 2],
        ["Ar-Ge Proje Raporu Taslağı", "Proje notları ve sonuçlarından TÜBİTAK/teşvik formatına uygun dönem raporu taslağı hazırlar.", 3],
      ]},
    ],
  },
};

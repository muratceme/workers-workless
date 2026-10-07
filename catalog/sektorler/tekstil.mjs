/* Tekstil ve Hazır Giyim (ihracatçı konfeksiyon odaklı) — unvanlar MRM Konfeksiyon, Alle Tekstil,
   Fateks, Hyosung, Beyyoğlu Tekstil, Efe Tekstil ilanlarına dayanır. */

export const TEKSTIL = {
  "tekstil-merchandising": {
    name: "Müşteri Temsilciliği (Merchandising)", short: "MRC",
    desc: "Numuneden sevkiyata sipariş takibi ve müşteri koordinasyonu.",
    roles: [
      { id: "development-merchandiser", name: "Development Merchandiser", level: "Uzman", tasks: [
        ["Numune Takip Çizelgesi", "Numune taleplerini (proto, fit, PP, SMS) aşama, gönderim tarihi ve müşteri yorumlarıyla takip eder, geciken numuneleri listeler.", 4],
        ["Ürün Maliyet Hesaplama (Costing)", "Kumaş tüketimi ve fiyatı, aksesuar, baskı-nakış, CM (kesim-dikim) ve genel giderlerden ürün FOB maliyeti ve teklif fiyatı hesaplar.", 4],
        ["Müşteri Yorumları (Fit Comments) Özeti", "Müşterinin fit ve numune yorum e-postalarını modelhane ve üretim için madde madde aksiyon listesine çevirir.", 3],
      ]},
      { id: "production-merchandiser", name: "Production Merchandiser", level: "Uzman", tasks: [
        ["Sipariş Termin Planı (T&A)", "Sipariş sevk tarihinden geriye doğru kumaş, aksesuar, kesim, dikim, ütü-paket ve final kontrol kilometre taşlarını hesaplar, gecikme riskini işaretler.", 5],
        ["Lab Dip ve Onay Takibi", "Renk (lab dip), baskı ve aksesuar onaylarının gönderim ve onay durumunu sipariş bazında takip eder.", 2],
        ["Müşteri Haftalık Durum Raporu", "Siparişlerin üretim aşaması ve sevk tarihi durumunu müşteriye gönderilecek haftalık rapor formatına çevirir.", 3],
      ]},
    ],
  },

  "tekstil-tasarim": {
    name: "Tasarım ve Koleksiyon", short: "TSR",
    desc: "Koleksiyon hazırlığı, trend, renk ve teknik föy.",
    roles: [
      { id: "moda-tasarimcisi", name: "Moda Tasarımcısı", level: "Uzman", tasks: [
        ["Renk Paleti Çıkarma", "Görsellerden baskın renkleri çıkarıp en yakın Pantone TCX kodlarıyla renk paleti oluşturur.", 2],
        ["Teknik Föy (Tech Pack) Taslağı", "Model bilgisi, ölçüler, kumaş ve aksesuar listesinden teknik föy taslağını hazırlar.", 5],
        ["Trend Raporu Özeti", "Sezon trend raporlarını koleksiyona uygulanabilir renk, kumaş ve detay maddelerine indirir.", 2],
      ]},
    ],
  },

  "tekstil-modelhane": {
    name: "Modelhane", short: "MDH",
    desc: "Kalıp, ölçü tablosu, beden serisi ve numune dikimi.",
    roles: [
      { id: "modelist", name: "Modelist", level: "Uzman", tasks: [
        ["Ölçü Tablosu Beden Serisi (Grading)", "Ana beden ölçülerinden beden artış kurallarıyla tüm beden serisinin ölçü tablosunu üretir.", 4],
        ["Numune Ölçü Kontrolü", "Numune ölçümlerini ölçü tablosu ve toleranslarla karşılaştırıp tolerans dışı noktaları raporlar.", 3],
      ]},
    ],
  },

  "tekstil-planlama": {
    name: "Planlama", short: "PLN",
    desc: "Kumaş-aksesuar ihtiyacı, kesim ve üretim planlaması.",
    roles: [
      { id: "tekstil-planlama-uzmani", name: "Planlama Uzmanı", level: "Uzman", tasks: [
        ["Kumaş ve Aksesuar İhtiyaç Hesabı", "Sipariş adetleri, beden-renk dağılımı, birim tüketim ve fire oranlarından kumaş metrajını ve aksesuar adetlerini hesaplar.", 5],
        ["Kesim Kat Planı", "Sipariş beden dağılımı, pastal boyu ve maksimum kat sayısına göre pastal (marker) kombinasyonlarını ve kat adetlerini çıkarır.", 4],
        ["Üretim Hattı Yükleme Planı", "Siparişlerin dakika değeri (SAM) ve hat kapasitesinden hat bazında yükleme planı ve tahmini bitiş tarihlerini hesaplar.", 4],
      ]},
    ],
  },

  "tekstil-tedarik": {
    name: "Kumaş ve Aksesuar Satın Alma", short: "TDR",
    desc: "Kumaş, iplik ve aksesuar tedariki.",
    roles: [
      { id: "kumas-satin-alma-uzmani", name: "Kumaş Satın Alma Uzmanı", level: "Uzman", tasks: [
        ["Kumaş Sipariş Termin Takibi", "Kumaş siparişlerinin boyahane/örme aşamalarını ve termin durumunu takip eder, sipariş sevk tarihine etkisini hesaplar.", 4],
        "Teklif Karşılaştırma",
        ["Kumaş Fiyat Takibi", "Tedarikçi kumaş ve iplik fiyat listelerini toplayıp kalite bazında fiyat geçmişi tablosu oluşturur.", 2],
      ]},
    ],
  },

  "tekstil-fason": {
    name: "Üretim ve Fason Takip", short: "FSN",
    desc: "Kesimhane, dikim ve fason atölye koordinasyonu.",
    roles: [
      { id: "kesimhane-sefi", name: "Kesimhane Şefi", level: "Uzman", tasks: [
        "Kesim Kat Planı",
        ["Kesim Föyü ve Fire Raporu", "Kesilen adetleri beden-renk bazında sipariş adetleriyle karşılaştırır, kumaş kullanımını ve kesim firesini hesaplar.", 3],
      ]},
      { id: "fason-takip-sorumlusu", name: "Fason Takip Sorumlusu", level: "Uzman", tasks: [
        ["Fason Atölye İş Takibi", "Fason atölyelere çıkan işleri adet, çıkış-dönüş tarihi ve kalite durumuyla takip eder, gecikmeleri listeler.", 4],
        ["Fason Hakediş Hesaplama", "Atölyeden dönen sağlam adetler, birim fiyat, eksik ve hatalı ürün kesintilerinden fason hakedişini hesaplar.", 3],
      ]},
    ],
  },

  "tekstil-kalite": {
    name: "Kalite Kontrol", short: "KLT",
    desc: "Kumaş kontrol, ara kontrol ve final (AQL) kontrol.",
    roles: [
      { id: "tekstil-kalite-kontrol-elemani", name: "Kalite Kontrol Elemanı", level: "Giriş", tasks: [
        ["Kumaş Kontrol (4 Puan Sistemi) Raporu", "Top bazında kumaş hata kayıtlarından 4 puan sistemine göre 100 yard² başına puanı hesaplar, kabul/ret durumunu raporlar.", 3],
        "AQL Örneklem Planı",
      ]},
      { id: "kalite-kontrol-sefi-tekstil", name: "Kalite Kontrol Şefi", level: "Yönetici", tasks: [
        ["Final Kontrol Raporu", "Final denetim kayıtlarından kritik, majör ve minör hata sayılarını AQL kabul sayılarıyla karşılaştırıp sonuç raporu üretir.", 3],
        ["Hata Türü Pareto Analizi", "Ara ve final kontrol hata kayıtlarını hata türü, operasyon ve hat bazında Pareto analizine tabi tutar.", 2],
      ]},
    ],
  },

  "tekstil-sevkiyat": {
    name: "Ütü-Paket ve Sevkiyat", short: "SVK",
    desc: "Paketleme, koli listesi ve sevkiyat evrakı.",
    roles: [
      { id: "sevkiyat-sorumlusu-tekstil", name: "Sevkiyat Sorumlusu", level: "Uzman", tasks: [
        ["Koli Listesi (Packing List) Hazırlama", "Sipariş beden-renk dağılımı ve koli kurallarından (solid/asorti) koli bazında çeki listesi, brüt-net ağırlık ve hacim hesaplar.", 4],
        "Proforma Fatura Hazırlama",
      ]},
    ],
  },

  "tekstil-uygunluk": {
    name: "Sürdürülebilirlik ve Sosyal Uygunluk", short: "UYG",
    desc: "Müşteri ve sosyal uygunluk denetimleri, sürdürülebilirlik raporlaması.",
    roles: [
      { id: "sosyal-uygunluk-uzmani", name: "Sosyal Uygunluk Uzmanı", level: "Uzman", tasks: [
        ["Denetim Bulgu ve Düzeltici Faaliyet Takibi", "BSCI, Sedex ve müşteri denetim bulgularını düzeltici faaliyet planı (CAP), sorumlu ve termin bazında takip eder.", 3],
        ["Karbon Ayak İzi Hesaplama", "Enerji, yakıt ve malzeme tüketim verisinden kapsam 1-2 emisyonlarını ve ürün başına karbon değerini hesaplar.", 3],
      ]},
    ],
  },
};

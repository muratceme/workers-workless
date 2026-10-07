/* Gıda Üretimi — unvanlar Acemoğlu Gıda, Tatbak, Detay Gıda, Gurme, Adova ilanlarına dayanır
   (HACCP, GMP, ISO 22000/BRC/IFS, DÖF, izlenebilirlik). */

export const GIDA = {
  "gida-kalite-guvence": {
    name: "Kalite Güvence ve Gıda Güvenliği", short: "KGV",
    desc: "HACCP, GMP, denetimler, izlenebilirlik ve DÖF.",
    roles: [
      { id: "kalite-guvence-uzmani-gida", name: "Kalite Güvence Uzmanı", level: "Uzman", tasks: [
        ["HACCP Kritik Kontrol Noktası İzleme Analizi", "KKN izleme kayıtlarını (sıcaklık, süre, pH, metal dedektör) kritik limitlerle karşılaştırıp sapmaları ve düzeltici faaliyet ihtiyacını listeler.", 4],
        ["Geri Çağırma (İzlenebilirlik) Simülasyonu", "Hammadde lot numarasından etkilenen üretim partilerini, sevk edilen müşterileri ve miktarları ileri-geri izlenebilirlikle çıkarır.", 3],
        "DÖF Takibi",
        ["Denetim Hazırlık Kontrol Listesi", "BRC, IFS veya resmî denetim maddelerine göre doküman ve kayıt hazırlık durumunu takip eden kontrol listesi üretir.", 2],
      ]},
      { id: "gida-muhendisi", name: "Gıda Mühendisi", level: "Uzman", tasks: [
        ["Etiket Bilgisi Kontrolü", "Ürün etiketini Türk Gıda Kodeksi etiketleme kurallarına (zorunlu bilgiler, alerjen vurgusu, net miktar) göre kontrol eder.", 3],
        ["Besin Değeri Hesaplama", "Reçetedeki bileşenlerin besin değeri tablosundan 100 g başına enerji ve besin öğelerini hesaplar.", 2],
      ]},
    ],
  },

  "gida-laboratuvar": {
    name: "Kalite Kontrol ve Laboratuvar", short: "LAB",
    desc: "Hammadde, ara ürün ve mamul analizleri.",
    roles: [
      { id: "kalite-kontrol-teknisyeni-gida", name: "Kalite Kontrol Teknisyeni", level: "Giriş", tasks: [
        ["Analiz Sonucu Spesifikasyon Kontrolü", "Laboratuvar analiz sonuçlarını ürün spesifikasyon limitleriyle karşılaştırıp uygunsuz partileri ve sınıra yakın değerleri işaretler.", 3],
        ["Raf Ömrü Takibi", "Stok ve sevkiyat verisinden son tüketim tarihine kalan süreyi hesaplar, müşteri kabul kriterini karşılamayan partileri listeler.", 2],
      ]},
    ],
  },

  "gida-uretim": {
    name: "Üretim", short: "ÜRT",
    desc: "Üretim partileri, verim, fire ve hijyen.",
    roles: [
      { id: "uretim-sefi-gida", name: "Üretim Şefi", level: "Yönetici", tasks: [
        ["Parti Verim ve Fire Raporu", "Parti bazında hammadde girişi, mamul çıkışı ve fireden verimi hesaplar, hedef dışı partileri listeler.", 3],
        "Vardiya Üretim Raporu",
        ["Temizlik ve Sanitasyon Kayıt Kontrolü", "Hat temizlik ve sanitasyon kayıtlarını plana göre kontrol edip eksik veya geciken kayıtları listeler.", 2],
      ]},
    ],
  },

  "gida-arge": {
    name: "Ar-Ge ve Ürün Geliştirme", short: "ARG",
    desc: "Yeni ürün, reçete ve maliyet çalışmaları.",
    roles: [
      { id: "ar-ge-uzmani-gida", name: "Ar-Ge Uzmanı", level: "Uzman", tasks: [
        ["Reçete Maliyet Hesaplama", "Reçete bileşen oranları, hammadde fiyatları ve fire oranlarından kilogram ve ambalaj başına maliyeti hesaplar.", 3],
        "Besin Değeri Hesaplama",
        ["Duyusal Test Sonuç Analizi", "Panel değerlendirme puanlarını numune ve özellik bazında istatistiksel olarak karşılaştırır.", 2],
      ]},
    ],
  },

  "gida-satis": {
    name: "Satış ve Kanal Yönetimi", short: "KNL",
    desc: "Ulusal zincir, geleneksel kanal ve bayi satışları.",
    roles: [
      { id: "kanal-satis-yoneticisi", name: "Kanal Satış Yöneticisi", level: "Yönetici", tasks: [
        ["Bayi ve Distribütör Satış Raporu", "Bayi bazında satış, hedef, tahsilat ve stok günlerini raporlar.", 3],
        ["Zincir Market Listeleme Takibi", "Zincir marketlerde listelenen ürünleri, raf fiyatlarını ve stoksuz mağazaları takip eder.", 2],
      ]},
    ],
  },
};

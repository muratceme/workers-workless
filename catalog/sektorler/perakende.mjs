/* Perakende (mağazacılık) — unvanlar zincir mağaza ilanları, Yargıcı "Alokasyon Uzmanı",
   Orka Holding "Ürün Yönetimi ve Planlama Uzmanı", eleman.net/secretcv rehberlerine dayanır. */

export const PERAKENDE = {
  "perakende-magaza": {
    name: "Mağaza Operasyonları", short: "MĞZ",
    desc: "Satış, kasa, reyon, vardiya ve mağaza performansı.",
    roles: [
      { id: "satis-danismani", name: "Satış Danışmanı", level: "Giriş", tasks: [
        ["Ürün Bilgi Kartı Özeti", "Ürün teknik bilgilerinden müşteriye anlatılacak kısa özellik ve fark özetleri hazırlar.", 1],
        ["Stok Sorgu ve Transfer Talebi", "Mağazada olmayan beden/rengin diğer mağaza stoklarını listeleyip transfer talebi formatına çevirir.", 1],
      ]},
      { id: "kasa-sorumlusu", name: "Kasa Sorumlusu", level: "Giriş", tasks: [
        ["Gün Sonu Kasa Raporu", "POS ve yazar kasa Z raporlarından nakit, kredi kartı, iade ve iptal toplamlarını çıkarır, kasa sayımıyla farkı raporlar.", 2],
        ["Kasa Fark ve İptal Analizi", "Kasiyer bazında iade, iptal ve kasa farklarını analiz edip olağandışı örüntüleri işaretler.", 2],
      ]},
      { id: "magaza-mudur-yardimcisi", name: "Mağaza Müdür Yardımcısı", level: "Uzman", tasks: [
        ["Mağaza Vardiya Planlama", "Saatlik müşteri trafiği tahmini, personel izinleri ve çalışma süresi kurallarına göre haftalık vardiya planı çıkarır.", 3],
        "Stok Sayım Fark Analizi",
        ["Fiyat Değişikliği Etiket Listesi", "Merkezden gelen fiyat değişikliklerini reyon ve raf sırasına göre etiket değişim listesine çevirir.", 1],
      ]},
      { id: "magaza-muduru", name: "Mağaza Müdürü", level: "Yönetici", tasks: [
        ["Mağaza Satış Performans Panosu", "Ciro, fiş adedi, sepet ortalaması, dönüşüm oranı ve geçen yıl karşılaştırmasını (LFL) günlük panoda toplar.", 2],
        ["Fire ve Kayıp Analizi", "Fire kayıtlarını ürün grubu ve neden bazında analiz eder, sayım farklarıyla birlikte kayıp oranını hesaplar.", 2],
        ["Personel Satış Performansı", "Satış danışmanlarının ciro, adet ve sepet ortalamalarını karşılaştırır, prim tablosunu hazırlar.", 2],
      ]},
      { id: "bolge-muduru-perakende", name: "Bölge Müdürü", level: "Yönetici", tasks: [
        ["Bölge Mağaza Karşılaştırma Raporu", "Bölgedeki mağazaları metrekare verimliliği, LFL büyüme, dönüşüm ve stok devir hızına göre karşılaştırır.", 2],
        ["Mağaza Ziyaret Raporu", "Mağaza ziyaret kontrol listesi notlarından aksiyon ve takip raporu hazırlar.", 2],
      ]},
    ],
  },

  "perakende-kategori": {
    name: "Satın Alma ve Kategori Yönetimi", short: "KTG",
    desc: "Ürün gamı, tedarikçi pazarlığı, fiyatlama ve kampanyalar.",
    roles: [
      { id: "kategori-uzmani", name: "Kategori Uzmanı", level: "Uzman", tasks: [
        ["Kategori Performans Raporu", "Ürün ve marka bazında satış, marj, stok devir hızı ve pazar payını raporlar; düşük performanslı ürünleri işaretler.", 3],
        ["Kampanya Etki Analizi", "Kampanya dönemindeki ek satışı (uplift), kanibalizasyonu ve kampanya kârlılığını hesaplar.", 3],
        ["Rakip Fiyat Karşılaştırması", "Rakip fiyat taramalarını ürün eşleştirmesiyle karşılaştırıp fiyat endeksi çıkarır.", 2],
      ]},
      { id: "kategori-yoneticisi", name: "Kategori Yöneticisi", level: "Yönetici", tasks: [
        ["Ürün Gamı (Delist) Önerisi", "Satış, marj ve raf verimliliğine göre listeden çıkarılabilecek ürünleri önerir.", 2],
        ["Tedarikçi Ciro Primi Hesabı", "Tedarikçi anlaşmalarındaki ciro primi ve iskonto kademelerine göre hak edilen primleri hesaplar.", 2],
      ]},
    ],
  },

  "perakende-planlama": {
    name: "Ürün Planlama ve Alokasyon", short: "ALK",
    desc: "Mağaza dağıtımı, otomatik sevkiyat, transfer ve stok planlama.",
    roles: [
      { id: "alokasyon-uzmani", name: "Alokasyon Uzmanı", level: "Uzman", tasks: [
        ["İlk Dağıtım (Alokasyon) Planı", "Yeni sezon ürünlerini mağaza kümesi, satış potansiyeli ve beden eğrisine göre mağazalara dağıtım adetlerine çevirir.", 5],
        ["Otomatik Sevkiyat (Replenishment) Önerisi", "Mağaza satış hızı, mevcut stok ve hedef stok gününe göre depodan sevk edilecek miktarları hesaplar.", 4],
        ["Mağazalar Arası Transfer Önerisi", "Fazla stoklu ve tükenme riskli mağazaları eşleştirip beden-renk bazında transfer önerisi çıkarır.", 3],
        ["Satış Hızı (Sell-Through) Raporu", "Ürün ve mağaza bazında satış oranı ve stok günü raporlar, indirime aday ürünleri belirler.", 2],
      ]},
    ],
  },

  "perakende-gorsel": {
    name: "Görsel Düzenleme", short: "GRS",
    desc: "Vitrin, raf düzeni ve planogram uyumu.",
    roles: [
      { id: "gorsel-duzenleme-uzmani", name: "Görsel Düzenleme Uzmanı", level: "Uzman", tasks: [
        ["Planogram Uyum Kontrolü", "Raf fotoğrafı ve planogram listesini karşılaştırıp eksik, yanlış yerleştirilmiş ve fiyat etiketi olmayan ürünleri listeler.", 3],
        ["Vitrin Yönergesi Hazırlama", "Kampanya ve ürün listesinden mağazalara gönderilecek adım adım vitrin düzenleme yönergesi hazırlar.", 2],
      ]},
    ],
  },

  "perakende-crm": {
    name: "CRM ve Sadakat Programı", short: "CRM",
    desc: "Müşteri verisi, sadakat programı ve kişiselleştirilmiş kampanyalar.",
    roles: [
      { id: "crm-uzmani", name: "CRM Uzmanı", level: "Uzman", tasks: [
        "Müşteri Segmentasyonu (RFM)",
        ["Sepet Birliktelik Analizi", "Fiş satırlarından birlikte alınan ürün çiftlerini destek, güven ve lift değerleriyle çıkarır.", 3],
        ["Kayıp Müşteri (Churn) Listesi", "Alışveriş sıklığı düşen sadakat programı üyelerini belirleyip geri kazanım kampanyası listesi çıkarır.", 2],
      ]},
    ],
  },

  "perakende-kayip-onleme": {
    name: "Kayıp Önleme", short: "KÖN",
    desc: "Envanter kaybı, kasa suistimali ve mağaza denetimleri.",
    roles: [
      { id: "kayip-onleme-uzmani", name: "Kayıp Önleme Uzmanı", level: "Uzman", tasks: [
        "Kasa Fark ve İptal Analizi",
        ["Envanter Kayıp Risk Skoru", "Sayım farkı, fire, iade ve iptal göstergelerinden mağaza bazında kayıp risk skoru üretir, denetim önceliği çıkarır.", 3],
      ]},
    ],
  },
};

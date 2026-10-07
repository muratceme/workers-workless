/* Turizm ve Otelcilik — departmanlar otel organizasyon yapısı (ön büro, kat hizmetleri, yiyecek-içecek,
   satış-pazarlama, gelir yönetimi); unvanlar Koru Hotels, Bomonti Suites, Rixos Pera ilanlarına dayanır. */

export const OTEL = {
  "otel-on-buro": {
    name: "Ön Büro", short: "ÖNB",
    desc: "Resepsiyon, giriş-çıkış işlemleri, gece denetimi ve misafir talepleri.",
    roles: [
      { id: "resepsiyonist", name: "Resepsiyonist", level: "Giriş", tasks: [
        ["Günlük Giriş-Çıkış Listesi", "Rezervasyonlardan günlük gelecek ve ayrılacak misafirleri oda tipi, özel istek ve ödeme durumuyla listeler.", 2],
        ["Misafir Mesajlarına Cevap Taslağı", "Misafir e-posta ve mesajlarına otel bilgilerine dayanarak çok dilli cevap taslağı hazırlar.", 3],
      ]},
      { id: "night-auditor", name: "Night Auditor (Gece Denetçisi)", level: "Uzman", tasks: [
        ["Gece Denetimi (Night Audit) Kontrolü", "Gün sonu folyo, tahsilat, oda geliri ve kasa raporlarını karşılaştırıp açık folyo, fiyat hatası ve kasa farklarını raporlar.", 4],
      ]},
      { id: "on-buro-muduru", name: "Ön Büro Müdürü", level: "Yönetici", tasks: [
        ["Doluluk, ADR ve RevPAR Raporu", "Oda satışlarından günlük, aylık ve geçen yıl karşılaştırmalı doluluk, ortalama oda fiyatı ve RevPAR'ı hesaplar.", 3],
        ["İptal ve No-Show Analizi", "Rezervasyon iptal ve gelmeme oranlarını kanal, oda tipi ve rezervasyon zamanlamasına göre analiz eder.", 2],
      ]},
    ],
  },

  "otel-rezervasyon": {
    name: "Rezervasyon ve Gelir Yönetimi", short: "GEL",
    desc: "Satış kanalları, fiyatlama ve doluluk optimizasyonu.",
    roles: [
      { id: "rezervasyon-sorumlusu", name: "Rezervasyon Sorumlusu", level: "Uzman", tasks: [
        ["Acente Kontenjan (Allotment) Takibi", "Tur operatörü kontratlarındaki kontenjan, release ve kullanım durumunu tarih bazında takip eder.", 3],
        ["Grup Rezervasyon Teklifi", "Grup talebinden oda, yemek ve toplantı salonu fiyatlarıyla teklif dosyası hazırlar.", 2],
      ]},
      { id: "gelir-muduru", name: "Gelir Müdürü (Revenue Manager)", level: "Yönetici", tasks: [
        ["Rezervasyon Pick-up Raporu", "Gelecek tarihler için günlük rezervasyon artışını (pick-up) geçen yıl ve bütçeyle karşılaştırır.", 3],
        ["Kanal Maliyet Analizi", "OTA, acente ve direkt kanallardan gelen gelirin komisyon sonrası net getirisini karşılaştırır.", 2],
        ["Dinamik Fiyat Önerisi", "Doluluk tahmini, pick-up hızı ve rakip fiyatlarına göre tarih ve oda tipi bazında fiyat önerisi üretir.", 3],
      ]},
    ],
  },

  "otel-kat-hizmetleri": {
    name: "Kat Hizmetleri", short: "KAT",
    desc: "Oda temizliği, kat planlaması ve çamaşırhane.",
    roles: [
      { id: "kat-hizmetleri-muduru", name: "Kat Hizmetleri Müdürü (Housekeeper)", level: "Yönetici", tasks: [
        ["Oda Temizlik Atama Planı", "Giriş-çıkış ve konaklama odalarını temizlik süresi kredilerine göre kat görevlilerine dengeli şekilde atar.", 3],
        ["Oda Arıza ve Bakım Talep Takibi", "Odalardan bildirilen arızaları oda, tür ve bekleme süresine göre takip eder.", 2],
      ]},
    ],
  },

  "otel-yiyecek-icecek": {
    name: "Yiyecek ve İçecek", short: "Yİ",
    desc: "Restoran, bar, mutfak maliyetleri ve menü.",
    roles: [
      { id: "yiyecek-icecek-muduru", name: "Yiyecek İçecek Müdürü", level: "Yönetici", tasks: [
        ["Menü Mühendisliği Analizi", "Ürün satış adetleri ve katkı paylarından menü kalemlerini yıldız, at, bulmaca ve köpek sınıflarına ayırır, aksiyon önerir.", 3],
        ["Porsiyon Maliyeti Hesaplama", "Standart reçete, porsiyon gramajı ve güncel alış fiyatlarından menü kalemi başına porsiyon maliyetini ve maliyet oranını hesaplar.", 3],
        ["Yiyecek-İçecek Maliyet Oranı Raporu", "Dönem tüketimi ve satışlardan departman bazında yiyecek ve içecek maliyet oranını hesaplar.", 2],
      ]},
    ],
  },

  "otel-misafir-iliskileri": {
    name: "Misafir İlişkileri", short: "MİS",
    desc: "Misafir memnuniyeti, yorumlar ve şikâyetler.",
    roles: [
      { id: "misafir-iliskileri-muduru", name: "Misafir İlişkileri Müdürü", level: "Yönetici", tasks: [
        ["Misafir Yorum Analizi", "Online platform yorumlarını konu (temizlik, yemek, personel, oda) ve duygu bazında analiz edip haftalık rapor çıkarır.", 3],
        ["Yorum Cevap Taslağı", "Misafir yorumlarına otelin üslubuna uygun, kişiselleştirilmiş cevap taslağı hazırlar.", 3],
      ]},
    ],
  },
};

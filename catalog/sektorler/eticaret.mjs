/* E-ticaret (marka / satıcı tarafı) — unvanlar "E-Ticaret Uzmanı", "E-Ticaret Operasyon Uzmanı",
   "Pazaryeri ve E-Ticaret Uzmanı", "E-Ticaret Pazaryeri Sorumlusu" ilanlarına dayanır. */

export const ETICARET = {
  "eticaret-pazaryeri": {
    name: "Pazaryeri Yönetimi", short: "PZY",
    desc: "Trendyol, Hepsiburada, Amazon, N11 gibi kanallarda ürün, fiyat ve kampanya yönetimi.",
    roles: [
      { id: "pazaryeri-uzmani", name: "Pazaryeri Uzmanı", level: "Uzman", tasks: [
        ["Pazaryeri Sipariş Kârlılık Hesabı", "Sipariş bazında satış fiyatı, pazaryeri komisyonu, kargo, hizmet bedeli, ürün maliyeti ve KDV'den net kârı hesaplar; zararına satılan ürünleri işaretler.", 5],
        ["Toplu Ürün Yükleme Şablonu Hazırlama", "Ürün ana veri dosyasını pazaryerinin toplu yükleme şablonuna (kategori, özellik, varyant, görsel alanları) dönüştürür, eksik zorunlu alanları listeler.", 5],
        ["Kanallar Arası Stok ve Fiyat Kontrolü", "Farklı satış kanallarındaki stok ve fiyatları merkezi listeyle karşılaştırıp uyumsuzlukları listeler.", 3],
        ["Kampanya Fiyat Simülasyonu", "Planlanan indirim ve kupon kurgularının ürün bazında kâr marjına etkisini hesaplar.", 2],
      ]},
    ],
  },

  "eticaret-icerik": {
    name: "Kategori ve İçerik Yönetimi", short: "İÇR",
    desc: "Ürün içerikleri, SEO ve kategori performansı.",
    roles: [
      { id: "eticaret-icerik-uzmani", name: "E-ticaret İçerik Uzmanı", level: "Uzman", tasks: [
        ["Ürün Açıklaması Üretme", "Ürün özellik tablosundan SEO uyumlu, kanal karakter sınırlarına uygun başlık ve açıklama metinleri üretir.", 8],
        ["Ürün Başlığı SEO Kontrolü", "Ürün başlıklarını uzunluk, marka-ürün tipi-özellik sırası ve yasaklı ifadeler açısından kontrol edip öneri sunar.", 2],
      ]},
      { id: "eticaret-kategori-yoneticisi", name: "Kategori Yöneticisi", level: "Yönetici", tasks: [
        ["E-ticaret Satış Performans Raporu", "Kanal, kategori ve ürün bazında ciro, adet, dönüşüm ve iade oranını haftalık raporlar.", 3],
        "Rakip Fiyat Karşılaştırması",
      ]},
    ],
  },

  "eticaret-operasyon": {
    name: "E-ticaret Operasyon ve Kargo", short: "OPR",
    desc: "Sipariş karşılama, kargo performansı ve iade süreçleri.",
    roles: [
      { id: "eticaret-operasyon-uzmani", name: "E-ticaret Operasyon Uzmanı", level: "Uzman", tasks: [
        ["Geciken Sipariş Takibi", "Kargoya verilmesi gereken süreyi aşan siparişleri kanal ve depo bazında listeler.", 3],
        ["Kargo Teslim Performansı", "Kargo firması bazında teslim süresi, hasarlı ve kayıp gönderi oranlarını hesaplar.", 2],
        ["İade Nedeni Analizi", "İade açıklamalarını neden kategorilerine (beden, kusur, beklenti farkı) ayırıp ürün bazında iade oranı çıkarır.", 3],
      ]},
    ],
  },

  "eticaret-performans": {
    name: "Performans Pazarlama", short: "PRF",
    desc: "Reklam yönetimi, dönüşüm optimizasyonu ve analitik.",
    roles: [
      { id: "performans-pazarlama-uzmani", name: "Performans Pazarlama Uzmanı", level: "Uzman", tasks: [
        "Reklam Kampanyası Performans Raporu",
        ["Pazaryeri Reklam (Sponsorlu Ürün) Analizi", "Pazaryeri reklam raporlarından ürün bazında reklam maliyeti/satış oranını hesaplayıp kârsız reklamları işaretler.", 3],
      ]},
    ],
  },

  "eticaret-musteri": {
    name: "Müşteri Deneyimi", short: "MDN",
    desc: "Ürün soruları, değerlendirmeler ve müşteri iletişimi.",
    roles: [
      { id: "eticaret-musteri-hizmetleri-temsilcisi", name: "Müşteri Hizmetleri Temsilcisi", level: "Giriş", tasks: [
        ["Ürün Sorusu Cevap Taslağı", "Pazaryeri ürün sorularına ürün bilgisi ve şirket politikalarına dayanarak cevap taslağı hazırlar.", 5],
        ["Ürün Değerlendirme Analizi", "Ürün yorumlarını duygu ve konu (kalite, beden, kargo, paketleme) bazında sınıflandırıp ürün bazında özetler.", 3],
      ]},
    ],
  },
};

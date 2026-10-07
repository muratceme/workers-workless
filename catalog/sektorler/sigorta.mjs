/* Sigortacılık — birim adları Anadolu Sigorta ve Ana Sigorta faaliyet raporlarına,
   unvanlar Türk Reasürans, Zurich, Şeker Sigorta, Anadolu Sigorta ilanlarına dayanır. */

export const SIGORTA = {
  "sigorta-teknik": {
    name: "Teknik (Risk Kabul / Underwriting)", short: "TKN",
    desc: "Teklif değerlendirme, risk kabul, fiyatlama ve poliçe üretimi.",
    roles: [
      { id: "teknik-uzman-yardimcisi", name: "Teknik Uzman Yardımcısı", level: "Giriş", tasks: [
        ["Teklif Talebi Bilgi Kontrolü", "Acenteden gelen teklif taleplerinde risk adresi, bedel, yapı tarzı, faaliyet konusu ve önceki hasar bilgisinin eksiksizliğini kontrol eder.", 3],
        ["Poliçe Üretim Kontrolü", "Üretilen poliçelerde teminat, bedel, prim ve özel şartları onaylı teklifle karşılaştırıp uyumsuzlukları listeler.", 3],
      ]},
      { id: "teknik-uzman", name: "Teknik Uzman (Underwriter)", level: "Uzman", tasks: [
        ["Risk Değerlendirme Özeti", "Risk inceleme (survey) raporundan yangın, deprem, hırsızlık ve sorumluluk risklerini, önerilen iyileştirmeleri ve kabul önerisini özetler.", 4],
        ["Sigorta Primi Hesaplama", "Tarife parametreleri, bedel, risk sınıfı ve indirim-sürprim kurallarıyla teminat bazında prim hesaplar.", 3],
        ["Hasar/Prim Oranı Analizi", "Poliçe ve hasar verisinden müşteri, acente ve branş bazında hasar/prim oranını hesaplayıp yenileme önerisi çıkarır.", 3],
      ]},
      { id: "teknik-mudur", name: "Teknik Müdür", level: "Yönetici", tasks: [
        ["Branş Kârlılık Raporu", "Branş bazında yazılan prim, kazanılmış prim, hasar ve giderlerden teknik kârlılık raporu üretir.", 2],
        ["Yetki Limiti Aşım Kontrolü", "Onaylanan tekliflerin teknik personel yetki limitlerine uygunluğunu kontrol eder.", 1],
      ]},
    ],
  },

  "sigorta-hasar": {
    name: "Hasar Yönetimi", short: "HSR",
    desc: "Hasar ihbarı, dosya yönetimi, eksper koordinasyonu ve rücu.",
    roles: [
      { id: "hasar-uzman-yardimcisi", name: "Hasar Uzman Yardımcısı", level: "Giriş", tasks: [
        ["Hasar Dosyası Evrak Eksik Kontrolü", "Branş ve hasar türüne göre gerekli evrak listesini dosyadaki belgelerle karşılaştırıp sigortalıdan istenecekleri çıkarır.", 4],
        ["Hasar İhbarı Ön Kayıt", "İhbar e-postası veya formundan poliçe no, hasar tarihi, yeri, türü ve tahmini tutarı ayıklayıp ön kayıt formatına çevirir.", 3],
      ]},
      { id: "hasar-uzmani", name: "Hasar Uzmanı", level: "Uzman", tasks: [
        ["Hasar Dosyası Özeti", "Ekspertiz raporu, tutanak, beyan ve faturalardan oluşan dosyayı karar için tek sayfalık özete indirir.", 6],
        ["Teminat Kapsam Kontrolü", "Hasar olayını poliçe teminatları, muafiyetler ve istisnalarla karşılaştırıp kapsam değerlendirmesi taslağı hazırlar.", 4],
        ["Suistimal Risk Göstergeleri", "Poliçe başlangıcına yakın hasar, sık hasar, tutarsız beyan ve şişirilmiş fatura gibi göstergelerden suistimal risk puanı üretir.", 3],
        ["Muallak Hasar Takibi", "Açık dosyaları yaş, rezerv tutarı ve bekleyen aksiyona göre takip eder, uzun süredir hareketsiz dosyaları listeler.", 2],
      ]},
      { id: "rucu-uzmani", name: "Rücu Uzmanı", level: "Uzman", tasks: [
        ["Rücu Potansiyeli Tespiti", "Ödenen hasarlarda kusur oranı, karşı taraf ve sigorta bilgisine göre rücu edilebilir dosyaları ve tutarlarını belirler.", 3],
        ["Rücu Tahsilat Takibi", "Rücu taleplerinin gönderim, cevap, tahkim ve tahsilat aşamalarını takip eder.", 2],
      ]},
      { id: "hasar-muduru", name: "Hasar Müdürü", level: "Yönetici", tasks: [
        ["Muallak Hasar Raporu", "Açık dosyaların branş, yaş ve rezerv tutarı dağılımını ve rezerv yeterliliği göstergelerini raporlar.", 2],
        ["Eksper Performans Analizi", "Eksperleri rapor süresi, tespit tutarı-ödeme farkı ve itiraz oranına göre karşılaştırır.", 2],
      ]},
    ],
  },

  "sigorta-saglik": {
    name: "Sağlık Sigortaları", short: "SĞL",
    desc: "Provizyon, sağlık hasar ödemeleri ve anlaşmalı kurum yönetimi.",
    roles: [
      { id: "provizyon-uzmani", name: "Provizyon Uzmanı", level: "Uzman", tasks: [
        ["Provizyon Talebi Değerlendirme", "Anlaşmalı kurumdan gelen provizyon talebini poliçe teminatı, limit, bekleme süresi ve ön mevcut durum istisnalarına göre ön değerlendirir.", 4],
        ["Sağlık Faturası Kontrolü", "Kurum faturalarını anlaşmalı fiyat listesi ve onaylı provizyonla satır bazında karşılaştırıp fazla tutarları işaretler.", 4],
      ]},
      { id: "anlasmali-kurumlar-uzmani-sigorta", name: "Anlaşmalı Kurumlar Uzmanı", level: "Uzman", tasks: [
        ["Kurum Fiyat Karşılaştırması", "Anlaşmalı sağlık kurumlarının işlem fiyatlarını bölge ve kurum grubu bazında karşılaştırır.", 2],
      ]},
    ],
  },

  "sigorta-akturya": {
    name: "Aktüerya", short: "AKT",
    desc: "Tarife, rezerv ve risk modelleri.",
    roles: [
      { id: "akturya-uzmani", name: "Aktüerya Uzmanı", level: "Uzman", tasks: [
        ["Hasar Sıklık ve Şiddet Analizi", "Poliçe ve hasar verisinden segment bazında hasar sıklığı, ortalama hasar ve risk primi hesaplar.", 4],
        ["IBNR Rezerv Tahmini (Zincirleme Merdiven)", "Hasar gelişim üçgeninden zincirleme merdiven yöntemiyle gerçekleşmiş ancak rapor edilmemiş hasar karşılığını hesaplar.", 4],
        ["Tarife Etki Simülasyonu", "Tarife parametresi değişikliğinin mevcut portföyde prim ve hasar/prim oranına etkisini simüle eder.", 3],
      ]},
    ],
  },

  "sigorta-reasurans": {
    name: "Reasürans", short: "RSR",
    desc: "Trete ve ihtiyari reasürans, reasürör hesapları.",
    roles: [
      { id: "reasurans-uzmani", name: "Reasürans Uzmanı", level: "Uzman", tasks: [
        ["Reasürans Hesap Cetveli Hazırlama", "Trete şartlarına göre devredilen prim, komisyon ve hasar paylarını dönem bazında hesaplayıp hesap cetvelini hazırlar.", 4],
        ["Kümül Risk Kontrolü", "Deprem bölgesi ve adres bazında aynı lokasyondaki poliçelerin toplam bedelini trete kapasitesiyle karşılaştırır.", 3],
      ]},
    ],
  },

  "sigorta-satis": {
    name: "Satış ve Acente Yönetimi", short: "ACN",
    desc: "Acente, broker ve banka kanallarının satış performansı.",
    roles: [
      { id: "bolge-satis-yoneticisi-sigorta", name: "Bölge Satış Yöneticisi", level: "Yönetici", tasks: [
        ["Acente Performans Karnesi", "Acente bazında üretim, büyüme, hasar/prim, tahsilat ve yenileme oranını karşılaştırmalı karnede toplar.", 3],
        ["Üretim Hedef Takibi", "Branş ve acente bazında aylık üretim hedefi-gerçekleşmesini takip eder.", 2],
      ]},
    ],
  },

  "sigorta-tahsilat": {
    name: "Tahsilat Yönetimi", short: "TAH",
    desc: "Prim tahsilatı, acente cari hesapları ve iptal süreçleri.",
    roles: [
      { id: "tahsilat-uzmani-sigorta", name: "Tahsilat Uzmanı", level: "Uzman", tasks: [
        ["Prim Alacak Yaşlandırma", "Acente ve poliçe bazında tahsil edilmemiş primleri vade gününe göre yaşlandırır.", 3],
        ["Prim Ödenmemesi Nedeniyle İptal Listesi", "Taksiti ödenmeyen poliçeleri genel şartlardaki süre kurallarına göre iptal adayı olarak listeler.", 2],
      ]},
    ],
  },
};

/* Sigorta acenteliği ve brokerlik — acente teknik/operasyon ilanlarına dayanır. */
export const ACENTE = {
  "acente-teknik": {
    name: "Teknik ve Operasyon", short: "TEK",
    desc: "Teklif alma, poliçe düzenleme, yenileme ve tahsilat.",
    roles: [
      { id: "sigorta-uzmani-acente", name: "Sigorta Uzmanı", level: "Uzman", tasks: [
        ["Çoklu Şirket Teklif Karşılaştırması", "Farklı sigorta şirketlerinden alınan teklifleri teminat, muafiyet, limit ve prim bazında müşteriye sunulacak tabloya döker.", 4],
        ["Poliçe Yenileme Takibi", "Vadesi yaklaşan poliçeleri listeler, müşteri için kişiselleştirilmiş yenileme hatırlatma mesajları hazırlar.", 3],
        ["Müşteri İhtiyaç Analizi", "Görüşme notlarından müşterinin risklerini ve önerilecek sigorta ürünlerini içeren ihtiyaç analizi hazırlar.", 2],
      ]},
      { id: "hasar-sorumlusu-acente", name: "Hasar Sorumlusu", level: "Uzman", tasks: [
        ["Müşteri Hasar Dosyası Takibi", "Müşterilerin açık hasar dosyalarını sigorta şirketi, eksper ve bekleyen evrak bazında takip eder.", 3],
        "Hasar Dosyası Evrak Eksik Kontrolü",
      ]},
      { id: "acente-yoneticisi", name: "Acente Yöneticisi", level: "Yönetici", tasks: [
        ["Acente Portföy ve Komisyon Raporu", "Şirket ve branş bazında üretim, komisyon geliri ve tahsilat durumunu raporlar.", 2],
        ["Müşteri Tahsilat Takibi", "Müşterilerden tahsil edilecek primleri vade ve şirkete aktarılacak tutarlarla birlikte takip eder.", 2],
      ]},
    ],
  },
};

/* Lojistik ve Taşımacılık (uluslararası nakliye, parsiyel, forwarding, depolama) — unvanlar
   MZM Global, Mars Lojistik, Evolog, Evrim Ulaştırma ve gümrük operasyon ilanlarına dayanır. */

export const LOJISTIK = {
  "lojistik-karayolu": {
    name: "Karayolu Operasyon", short: "KRY",
    desc: "Uluslararası ve yurt içi komple araç taşımaları.",
    roles: [
      { id: "karayolu-operasyon-uzmani", name: "Karayolu Operasyon Uzmanı", level: "Uzman", tasks: [
        ["Yük Planlama (LDM ve Palet)", "Yük ölçüleri, palet sayısı ve istiflenebilirliğe göre yükleme metresini (LDM), ağırlığı ve araç tipini hesaplar.", 3],
        ["Araç ve Sefer Takip Listesi", "Seferlerin yükleme, gümrük, sınır ve teslim aşamalarını takip eder, geciken araçları listeler.", 4],
        ["Müşteri Yük Durum Bilgilendirmesi", "Sefer durumundan müşteriye gönderilecek yük durum e-postalarını hazırlar.", 3],
        ["Sefer Kârlılık Hesabı", "Navlun geliri, alt nakliyeci bedeli, otoyol, gümrük ve diğer masraflardan sefer bazında kârlılığı hesaplar.", 3],
      ]},
    ],
  },

  "lojistik-parsiyel": {
    name: "Parsiyel Operasyon", short: "PRS",
    desc: "Grupaj (parsiyel) yük toplama, konsolidasyon ve dağıtım.",
    roles: [
      { id: "parsiyel-operasyon-uzmani", name: "Parsiyel Operasyon Uzmanı", level: "Uzman", tasks: [
        ["Parsiyel Yük Konsolidasyon Planı", "Bekleyen parsiyel yükleri varış bölgesi, hacim ve ağırlığa göre araçlara gruplayıp doluluk oranını hesaplar.", 4],
        ["Parsiyel Hakediş ve Kârlılık Hesabı", "Araçtaki yüklerin navlun gelirlerini, araç maliyetini ve yük bazında kâr payını hesaplar.", 3],
      ]},
    ],
  },

  "lojistik-forwarding": {
    name: "Deniz ve Hava Yolu Operasyon", short: "FWD",
    desc: "Konteyner, hava kargo rezervasyonu ve evrak süreçleri.",
    roles: [
      { id: "denizyolu-operasyon-uzmani", name: "Denizyolu Operasyon Uzmanı", level: "Uzman", tasks: [
        ["Navlun Teklif Hesaplama", "Ölçü ve ağırlıktan hava kargoda ücrete esas ağırlığı (1:6000), denizyolunda W/M'yi hesaplayıp navlun ve ek masraflarla teklif oluşturur.", 3],
        ["Konteyner Demuraj ve Detention Takibi", "Konteynerlerin limanda bekleme ve iade sürelerini serbest gün kurallarıyla izleyip oluşacak masrafı hesaplar.", 3],
        ["Konşimento Bilgi Çıkarma", "Konşimento (B/L) belgesinden yükleyici, alıcı, liman, konteyner ve yük bilgilerini yapılandırılmış kayda dönüştürür.", 3],
      ]},
    ],
  },

  "lojistik-gumruk": {
    name: "Gümrük Operasyon", short: "GMR",
    desc: "İthalat ve ihracat gümrük işlemleri.",
    roles: [
      { id: "gumruk-operasyon-uzmani", name: "Gümrük Operasyon Uzmanı", level: "Uzman", tasks: [
        ["Gümrük Evrak Tutarlılık Kontrolü", "Fatura, çeki listesi, menşe belgesi ve taşıma belgesindeki ürün, miktar, ağırlık ve değer bilgilerini karşılaştırıp uyumsuzlukları listeler.", 4],
        "GTİP Sınıflandırma Önerisi",
        ["İthalat Vergi Hesaplama", "CIF kıymet, gümrük vergisi, ilave gümrük vergisi ve KDV oranlarından ithalatta ödenecek vergileri hesaplar.", 3],
      ]},
    ],
  },

  "lojistik-depo": {
    name: "Depo ve Antrepo Yönetimi", short: "DPO",
    desc: "Depolama, elleçleme ve stok hizmeti verilen müşteriler.",
    roles: [
      { id: "depo-sefi", name: "Depo Şefi", level: "Uzman", tasks: [
        "Stok Sayım Fark Analizi",
        ["Depolama Hizmeti Faturalandırma", "Müşteri bazında palet-gün, giriş-çıkış elleçleme ve katma değerli hizmetlerden aylık depolama faturasını hesaplar.", 3],
        "Depo Yerleşim (ABC) Önerisi",
      ]},
    ],
  },

  "lojistik-filo": {
    name: "Filo Yönetimi", short: "FLO",
    desc: "Öz mal araçların bakım, yakıt ve sürücü yönetimi.",
    roles: [
      { id: "filo-yoneticisi", name: "Filo Yöneticisi", level: "Yönetici", tasks: [
        ["Yakıt Tüketim Analizi", "Yakıt kartı ve kilometre verisinden araç ve sürücü bazında 100 km tüketimini hesaplayıp sapmaları işaretler.", 3],
        ["Sürücü Çalışma ve Dinlenme Süresi Kontrolü", "Takograf verisinden günlük ve haftalık sürüş ile dinlenme sürelerini AETR kurallarına göre kontrol eder, ihlalleri listeler.", 3],
        "Araç Filosu Takibi",
      ]},
    ],
  },

  "lojistik-satis": {
    name: "Satış ve Fiyatlandırma", short: "SFY",
    desc: "Lojistik hizmet satışı, teklif ve müşteri yönetimi.",
    roles: [
      { id: "lojistik-satis-uzmani", name: "Lojistik Satış Uzmanı", level: "Uzman", tasks: [
        "Navlun Teklif Hesaplama",
        ["Teklif Dönüşüm Analizi", "Verilen navlun tekliflerinin kazanılma oranını hat, müşteri ve fiyat farkı bazında analiz eder.", 2],
      ]},
    ],
  },
};
